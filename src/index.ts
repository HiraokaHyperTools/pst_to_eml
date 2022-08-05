import { PSTContact, PSTMessage, PSTRecipient } from '@hiraokahypertools/pst-extractor';
import { PSTFile } from '@hiraokahypertools/pst-extractor';
import { PSTFolder } from '@hiraokahypertools/pst-extractor';
import MailComposer from 'nodemailer/lib/mail-composer';
import { Buffer } from 'buffer';
import { applyFallbackRecipients, changeFileExtension, convertToBuffer, formatFrom } from './utils';
import { convertVLines } from './vLines';

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
  toEmlBuffer(options: MsgConverterOptions): Promise<Buffer>;

  /**
   * Assume this is a contact and then convert to vCard.
   * 
   * @returns vCard text
   */
  toVCardStr(options: MsgConverterOptions): Promise<string>;
}

export interface IPFolder extends IPNode {
  subFolders(): Promise<PFolder[]>;
  items(): Promise<PItem[]>;
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

  async items(): Promise<PItem[]> {
    const list: PItem[] = [];

    if (this.folder.contentCount > 0) {
      const emails: PSTMessage[] = (await this.folder.getEmails());
      for (let email of emails) {
        list.push(new PItem(email));
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

export class PItem implements IPItem {
  private email: PSTMessage;

  async displayName(): Promise<string> {
    return this.email.subject || this.email.displayName;
  }

  constructor(email: PSTMessage) {
    this.email = email;
  }

  /**
   * Check `.msg` file usage
   * 
   * @returns
   * - This will return `IPM.Contact` for contact.
   * - This will return `IPM.Note` for EML.
   */
  public get messageClass(): string { return this.email.messageClass; }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML text
   */
  async toEmlStr(options: MsgConverterOptions): Promise<string> {
    return (await this.toEmlBuffer(options)).toString('utf-8');
  }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML file
   */
  async toEmlBuffer(options: MsgConverterOptions): Promise<Buffer> {
    return await this.toEmlFrom(options, this.email);
  }

  private async toEmlFrom(options: MsgConverterOptions, email: PSTMessage): Promise<Buffer> {
    const recipients = [];
    for (let x = 0; x < (await email.getNumberOfRecipients()); x++) {
      const entry: PSTRecipient = (await email.getRecipient(x));
      recipients.push({
        name: entry.displayName,
        email: entry.emailAddress,
        recipType: entry.recipientType,
      })
    }

    const attachmentsRefined = [];

    const entity = {
      baseBoundary: options.baseBoundary,

      from: formatFrom(email.senderName, email.senderEmailAddress),
      to: applyFallbackRecipients(
        recipients
          .map(
            ({ name, email, recipType }) => {
              return recipType === "to" ? { name, email } : null
            }
          )
          .filter((entry) => entry !== null), { name: "undisclosed-recipients" }),
      cc: recipients
        .map(({ name, email, recipType }) =>
          recipType === "cc" ? { name, email } : null
        )
        .filter((entry) => entry !== null),
      bcc: recipients
        .map(({ name, email, recipType }) =>
          recipType === "bcc" ? { name, email } : null
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
        const emlBuf = await this.toEmlFrom(options, embedded);

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
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML text
   */
  async toVCardStr(options: MsgConverterOptions): Promise<string> {
    return (await this.toVCardStrFrom(
      options,
      this.email as PSTContact
    ));
  }

  private async toVCardStrFrom(options: MsgConverterOptions, source: PSTContact): Promise<string> {

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
}
