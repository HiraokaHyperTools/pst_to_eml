import { PSTContact, PSTMessage, PSTRecipient } from '@hiraokahypertools/pst-extractor';
import { PSTFile } from '@hiraokahypertools/pst-extractor';
import { PSTFolder } from '@hiraokahypertools/pst-extractor';
import MailComposer from 'nodemailer/lib/mail-composer';
import { applyFallbackRecipients, changeFileExtension, convertToBuffer, formatAddress, formatFrom } from './utils';
import { convertVLines } from './vLines';
import { FasterEmail } from '@hiraokahypertools/pst-extractor/dist/FasterEmail';
import { decode } from 'iconv-lite';
import { Buffer } from 'buffer';

export async function wrapPstFile(
  pstFile: PSTFile
): Promise<PRoot> {
  if (!(pstFile instanceof PSTFile)) {
    console.warn("pstFile seems not to be instanceof PSTFile.");
  }
  return new PRoot(pstFile);
}

export interface IPNode {
  displayName(): Promise<string>;
}

export interface IPItem extends IPNode {
  /**
   * Check `.msg` file usage
   * 
   * @returns
   * - This will return `IPM.Contact` for contact.
   * - This will return `IPM.Note` for EML.
   */
  get messageClass(): string;

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML text
   */
  toEmlStr(options: MsgConverterOptions): Promise<string>;

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML file
   */
  toEmlBuffer(options: MsgConverterOptions): Promise<Uint8Array>;

  /**
   * Assume this is a contact and then convert to vCard.
   * 
   * @returns vCard text
   */
  toVCardStr(options: MsgConverterOptions): Promise<string>;
}

export interface IPFolder extends IPNode {
  subFolders(): Promise<PFolder[]>;
  items(options?: FolderItemsOptions): Promise<PItem[]>;
}

export interface IPRoot extends IPFolder {
  close(): Promise<void>;
}

export class PRoot implements IPRoot {
  private pstFile: PSTFile;
  private closed: boolean;

  constructor(pstFile: PSTFile) {
    this.pstFile = pstFile;
    this.closed = false;
  }

  async displayName(): Promise<string> {
    if (this.closed) {
      throw new Error("Trying to access disposed object.");
    }

    return (await this.pstFile.getMessageStore()).displayName;
  }

  async subFolders(): Promise<PFolder[]> {
    if (this.closed) {
      throw new Error("Trying to access disposed object.");
    }

    return [
      new PFolder((await this.pstFile.getRootFolder()))
    ]
  }

  async items(): Promise<PItem[]> {
    if (this.closed) {
      throw new Error("Trying to access disposed object.");
    }

    return []
  }

  async close(): Promise<void> {
    if (this.closed) {
      return;
    }

    this.pstFile.close();
    this.closed = true;
  }
}

export interface FolderItemsOptions {
  progress?: (current: number, count: number) => void;
}

export class PFolder implements IPFolder {
  private folder: PSTFolder;

  constructor(folder: PSTFolder) {
    this.folder = folder;
  }

  async displayName(): Promise<string> {
    return this.folder.displayName;
  }

  async subFolders(): Promise<PFolder[]> {
    const list: PFolder[] = [];

    if (this.folder.hasSubfolders) {
      const childFolders: PSTFolder[] = (await this.folder.getSubFolders());
      for (let childFolder of childFolders) {
        list.push(new PFolder(childFolder));
      }
    }

    return list;
  }

  async items(options?: FolderItemsOptions): Promise<PItem[]> {
    const list: PItem[] = [];

    if (1 <= this.folder.contentCount) {
      const fasterList = await this.folder.getFasterEmailList({
        progress: options?.progress,
      });

      for (let faster of fasterList) {
        list.push(new PItem(faster));
      }
    }

    return list
  }
}

export interface MsgConverterOptions {
  /**
   * fixed baseBoundary for EML for testing
   */
  baseBoundary?: string;

  /**
   * fixed messageId for EML for testing
   */
  messageId?: string;

}

const MAPI_TO = 1;
const MAPI_CC = 2;
const MAPI_BCC = 3;

export class PItem implements IPItem {
  private faster: FasterEmail;

  async displayName(): Promise<string> {
    return this.faster.displayName;
  }

  constructor(email: FasterEmail) {
    this.faster = email;
  }

  /**
   * Check `.msg` file usage
   * 
   * @returns
   * - This will return `IPM.Contact` for contact.
   * - This will return `IPM.Note` for EML.
   */
  public get messageClass(): string { return this.faster.messageClass; }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML text
   */
  async toEmlStr(options: MsgConverterOptions): Promise<string> {
    const buffer = await this.toEmlBuffer(options);
    return decode(Buffer.from(buffer), 'utf-8');
  }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML file
   */
  async toEmlBuffer(options: MsgConverterOptions): Promise<Uint8Array> {
    return await this.toEmlFrom(options, await this.faster.getMessage());
  }

  private async toEmlFrom(options: MsgConverterOptions, email: PSTMessage): Promise<Uint8Array> {
    return await toEmlFrom(options, email);
  }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML text
   */
  async toVCardStr(options: MsgConverterOptions): Promise<string> {
    return (await this.toVCardStrFrom(
      options,
      (await this.faster.getMessage()) as PSTContact
    ));
  }

  private async toVCardStrFrom(options: MsgConverterOptions, source: PSTContact): Promise<string> {
    return await toVCardStrFrom(options, source);
  }
}

/**
 * Convert a PST message to EML format.
 * 
 * @param options The conversion options.
 * @param email The PST message to convert.
 * @returns The EML data.
 */
export async function toEmlFrom(options: MsgConverterOptions, email: PSTMessage): Promise<Uint8Array> {
  const recipients = [];
  for (let x = 0; x < (await email.getNumberOfRecipients()); x++) {
    const entry: PSTRecipient = (await email.getRecipient(x));
    recipients.push({
      name: entry.displayName,
      email: entry.emailAddress,
      recipType: entry.recipientType,
    });
  }

  const attachmentsRefined = [];

  const entity = {
    baseBoundary: options.baseBoundary,

    from: formatFrom(email.senderName, email.senderEmailAddress),
    to: applyFallbackRecipients(
      recipients
        .map(
          ({ name, email, recipType }) => {
            return recipType === MAPI_TO ? formatAddress(name, email) : null
          }
        )
        .filter((entry) => entry !== null), { name: "undisclosed-recipients" }),
    cc: recipients
      .map(({ name, email, recipType }) =>
        recipType === MAPI_CC ? formatAddress(name, email) : null
      )
      .filter((entry) => entry !== null),
    bcc: recipients
      .map(({ name, email, recipType }) =>
        recipType === MAPI_BCC ? formatAddress(name, email) : null
      )
      .filter((entry) => entry !== null),
    subject: email.subject,
    text: email.body,
    html: email.bodyHTML,
    attachments: attachmentsRefined,
    headers: {
      "Date": (false
        || email.messageDeliveryTime
        || email.clientSubmitTime
        || email.modificationTime
        || email.creationTime
        || new Date()
      ).toString(),
      "Message-ID": options.messageId,
    }
  };

  for (let x = 0; x < (await email.getNumberOfAttachments()); x++) {
    const attachment = (await email.getAttachment(x));

    const embedded = (await attachment.getEmbeddedPSTMessage())
    if (embedded != null) {
      const emlBuf = await toEmlFrom(options, embedded);

      attachmentsRefined.push({
        filename: changeFileExtension(attachment.displayName ?? "unnamed", ".eml"),
        content: emlBuf,
        cid: attachment.contentId,
        contentTransferEncoding: '8bit',
      })
    }
    else {
      attachmentsRefined.push({
        filename: attachment.displayName,
        content: convertToBuffer(attachment.fileData),
        cid: attachment.contentId,
      });
    }
  }

  return await new Promise((resolve, reject) => {
    try {
      const mail = new MailComposer(entity);
      mail.compile().build(function (error, message) {
        if (error) {
          reject(new Error("EML composition failed.\n" + error + "\n\n" + error.stack));
          return;
        }
        else {
          resolve(message);
        }
      });
    } catch (ex) {
      reject(new Error("EML composition failed.\n" + ex));
    }
  });
}

/**
 * Convert a PST contact to a vCard string.
 * 
 * @param options The conversion options.
 * @param source The PST contact to convert.
 * @returns The vCard string.
 */
export async function toVCardStrFrom(options: MsgConverterOptions, source: PSTContact): Promise<string> {
  const makers = [
    {
      kind: "N",
      provider: () => [
        source.surname,
        source.givenName,
        source.middleName,
        source.displayNamePrefix,
        source.generation,
      ],
    }, {
      kind: "FN",
      provider: () => source.displayName,
    }, {
      kind: "X-MS-N-YOMI",
      provider: () => [source.yomiLastName, source.yomiFirstName],
    }, {
      kind: "ORG",
      provider: () => [source.companyName, source.departmentName],
    }, {
      kind: "X-MS-ORG-YOMI",
      provider: () => source.yomiCompanyName,
    }, {
      kind: "TITLE",
      provider: () => source.title,
    }, {
      kind: ["TEL", "WORK", "VOICE"],
      provider: () => source.businessTelephoneNumber,
    }, {
      kind: ["TEL", "HOME", "VOICE"],
      provider: () => source.homeTelephoneNumber,
    }, {
      kind: ["TEL", "CELL", "VOICE"],
      provider: () => source.mobileTelephoneNumber,
    }, {
      kind: ["TEL", "WORK", "FAX"],
      provider: () => source.businessFaxNumber,
    }, {
      kind: ["ADR", "WORK", "PREF"],
      provider: () => [
        source.workAddressStreet,
        source.workAddressCity,
        source.workAddressState,
        source.workAddressPostalCode,
        source.workAddressCountry
      ]
    }, {
      kind: ["LABEL", "WORK", "PREF"],
      provider: () => source.workAddress,
    }, {
      kind: ["URL", "WORK"],
      provider: () => source.businessHomePage,
    }, {
      kind: ["EMAIL", "PREF", "INTERNET"],
      provider: () => source.email1EmailAddress,
    }
  ];

  const vLines: [string | string[], string | string[]][] = [];
  vLines.push([["BEGIN"], "VCARD"]);
  vLines.push([["VERSION"], "2.1"]);

  makers.forEach(
    maker => {
      vLines.push([
        maker.kind,
        maker.provider()
      ]);
    }
  )

  vLines.push([["END"], "VCARD"]);

  return convertVLines(vLines);
}
