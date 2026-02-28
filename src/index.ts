import type { IPSTContact, IPSTMessage, IPSTRecipient, IPSTFile, IPSTFolder } from '@hiraokahypertools/pst-extractor';
import { applyFallbackRecipients, changeFileExtension, convertToUint8Array } from './utils.js';
import { convertVLines } from './vLines.js';
import { FasterEmail } from '@hiraokahypertools/pst-extractor';
import { Base64TransferEncoding, EmlWriter, EncodeWordInBase64 } from './EmlWriter.js';
import { v4 as uuidv4 } from 'uuid';

const utf8Decoder = new TextDecoder('utf-8');
const utf8Encoder = new TextEncoder();

interface AttachmentRefined {
  filename: string;
  content?: Uint8Array;
  nestedMail?: string;
  cid?: string;
}

export async function wrapPstFile(
  pstFile: IPSTFile
): Promise<PRoot> {
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

  get primaryNodeId(): number;

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
  get primaryNodeId(): number;
}

export interface IPRoot extends IPFolder {
  close(): Promise<void>;
}

export class PRoot implements IPRoot {
  private pstFile: IPSTFile;
  private closed: boolean;

  constructor(pstFile: IPSTFile) {
    this.pstFile = pstFile;
    this.closed = false;
  }

  async displayName(): Promise<string> {
    if (this.closed) {
      throw new Error("Trying to access disposed object.");
    }

    return (await this.pstFile.getMessageStore()).displayName;
  }

  get primaryNodeId(): number {
    if (this.closed) {
      throw new Error("Trying to access disposed object.");
    }

    return 0;
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
  private folder: IPSTFolder;

  constructor(folder: IPSTFolder) {
    this.folder = folder;
  }

  async displayName(): Promise<string> {
    return this.folder.displayName;
  }

  get primaryNodeId(): number {
    return this.folder.primaryNodeId;
  }

  async subFolders(): Promise<PFolder[]> {
    const list: PFolder[] = [];

    if (this.folder.hasSubfolders) {
      const childFolders: IPSTFolder[] = (await this.folder.getSubFolders());
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
  * fixed altBoundary for EML for testing
  */
  altBoundary?: string;

  /**
   * fixed messageId for EML for testing
   */
  messageId?: string;

  /**
   * Whether to allow nested EML. If this is false, nested EML will be treated as attachments with .eml extension. This is because some email clients do not support nested EML.
   */
  allowNestedEml?: boolean;
}

const MAPI_TO = 1;
const MAPI_CC = 2;
const MAPI_BCC = 3;

export class PItem implements IPItem {
  private faster: FasterEmail;

  async displayName(): Promise<string> {
    return this.faster.displayName;
  }

  get primaryNodeId(): number {
    return this.faster.primaryNodeId;
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
    return utf8Decoder.decode(buffer);
  }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML file
   */
  async toEmlBuffer(options: MsgConverterOptions): Promise<Uint8Array> {
    return await this.toEmlFrom(options, await this.faster.getMessage());
  }

  private async toEmlFrom(options: MsgConverterOptions, email: IPSTMessage): Promise<Uint8Array> {
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
      (await this.faster.getMessage()) as unknown as IPSTContact
    ));
  }

  private async toVCardStrFrom(options: MsgConverterOptions, source: IPSTContact): Promise<string> {
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
export async function toEmlFrom(options: MsgConverterOptions, email: IPSTMessage): Promise<Uint8Array> {
  const eml = await toEmlStringFrom(options, email);
  return utf8Encoder.encode(eml);
}

/**
 * Convert a PST message to EML format.
 * 
 * @param options The conversion options.
 * @param email The PST message to convert.
 * @returns The EML data.
 */
export async function toEmlStringFrom(options: MsgConverterOptions, email: IPSTMessage): Promise<string> {
  const recipients = [];
  for (let x = 0; x < (await email.getNumberOfRecipients()); x++) {
    const entry: IPSTRecipient = (await email.getRecipient(x));
    recipients.push({
      name: entry.displayName,
      email: entry.emailAddress,
      recipType: entry.recipientType,
    });
  }

  const attachmentsRefined: AttachmentRefined[] = [];

  for (let x = 0; x < (await email.getNumberOfAttachments()); x++) {
    const attachment = (await email.getAttachment(x));

    const embedded = (await attachment.getEmbeddedPSTMessage()) as IPSTMessage;
    const filename = [
      attachment.longFilename,
      attachment.filename,
      attachment.displayName,
      "unnamed"
    ]
      .filter(it => it)[0];
    if (embedded != null) {
      if (options.allowNestedEml) {
        const emlBuf = await toEmlStringFrom(options, embedded);

        attachmentsRefined.push({
          filename: changeFileExtension(filename, ".eml"),
          nestedMail: emlBuf,
          cid: attachment.contentId,
        });
      }
      else {
        const emlBuf = await toEmlFrom(options, embedded);

        attachmentsRefined.push({
          filename: changeFileExtension(filename, ".eml"),
          content: emlBuf,
          cid: attachment.contentId,
        });
      }
    }
    else {
      attachmentsRefined.push({
        filename: filename,
        content: convertToUint8Array(attachment.fileData),
        cid: attachment.contentId,
      });
    }
  }

  const boundary = uuidv4();
  const topBoundary = options.baseBoundary ?? ("_b1" + boundary + "_");
  const altBoundary = options.altBoundary ?? ("_b2" + boundary + "_");

  const base64Enc = new Base64TransferEncoding();
  const encodeWordInBase64 = new EncodeWordInBase64();

  const chunks: string[] = [];
  const emlWriter = new EmlWriter(
    it => chunks.push(it),
    word => encodeWordInBase64.encodeIfNeeded(word)
  );
  emlWriter
    .from({ name: email.senderName, email: email.senderEmailAddress })
    .to(
      applyFallbackRecipients(
        recipients
          .map(
            ({ name, email, recipType }) => {
              return recipType === MAPI_TO ? { name, email } : null
            }
          )
          .filter((entry) => entry !== null), { name: "undisclosed-recipients" })
    )
    .cc(
      recipients
        .map(({ name, email, recipType }) =>
          recipType === MAPI_CC ? { name, email } : null
        )
        .filter((entry) => entry !== null)
    )
    .bcc(
      recipients
        .map(({ name, email, recipType }) =>
          recipType === MAPI_BCC ? { name, email } : null
        )
        .filter((entry) => entry !== null)
    )
    .subject(email.subject)
    .date(
      (false
        || email.messageDeliveryTime
        || email.clientSubmitTime
        || email.modificationTime
        || email.creationTime
        || new Date()
      ).toString()
    )
    .messageId(options.messageId)
    .contentTypeMultipartMixed(topBoundary)
    .mimeVersion1()
    ;

  if (email.bodyHTML || email.body) {
    // multipart/alternative
    emlWriter
      .newLine()
      .beginBoundary(topBoundary)
      .contentTypeMultipartMixed(altBoundary);
    if (email.bodyHTML) {
      emlWriter
        .newLine()
        .beginBoundary(altBoundary)
        .contentType("text/html; charset=utf-8")
        .contentTransferEncoding(base64Enc.name)
        .newLine()
        .writeContent(base64Enc.applyStringAsUtf8(email.bodyHTML));
    }
    if (email.body) {
      emlWriter
        .newLine()
        .beginBoundary(altBoundary)
        .contentType("text/plain; charset=utf-8")
        .contentTransferEncoding(base64Enc.name)
        .newLine()
        .writeContent(base64Enc.applyStringAsUtf8(email.body));
    }
    emlWriter
      .newLine()
      .endBoundary(altBoundary);
  }

  for (const attachment of attachmentsRefined) {
    if (attachment.nestedMail) {
      emlWriter
        .newLine()
        .beginBoundary(topBoundary)
        .contentType(`message/rfc822; name=\"${encodeWordInBase64.encodeIfNeeded(attachment.filename)}\"`)
        .contentTransferEncoding("7bit");
      if (attachment.cid) {
        emlWriter
          .contentId(attachment.cid);
      }
      emlWriter
        .newLine();

      emlWriter.writeChunks(attachment.nestedMail.split(/\r?\n/).map(line => line + "\r\n"));
    }
    else if (attachment.content) {
      emlWriter
        .newLine()
        .beginBoundary(topBoundary)
        .contentType(`application/octet-stream; name=\"${encodeWordInBase64.encodeIfNeeded(attachment.filename)}\"`)
        .writeHeader("Content-Disposition", `attachment; filename="${encodeWordInBase64.encodeIfNeeded(attachment.filename)}"`)
        .contentTransferEncoding(base64Enc.name);
      if (attachment.cid) {
        emlWriter
          .contentId(attachment.cid);
      }
      emlWriter
        .newLine();

      emlWriter.writeContent(base64Enc.applyUint8Array(attachment.content));
    }
  }

  emlWriter
    .endBoundary(topBoundary);

  const eml = chunks.join("");
  return eml;
}

/**
 * Convert a PST contact to a vCard string.
 * 
 * @param options The conversion options.
 * @param source The PST contact to convert.
 * @returns The vCard string.
 */
export async function toVCardStrFrom(options: MsgConverterOptions, source: IPSTContact): Promise<string> {
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
