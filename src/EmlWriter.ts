interface HeaderWriter {
  value(rawString: string): HeaderWriter;
  end(): void;
}

type AddressSpec = { name: string; email?: string };

const utf8Encoder = new TextEncoder();

const base64Table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

export class Base64TransferEncoding {
  get name(): string {
    return "base64";
  }

  applyStringAsUtf8(data: string, cutLinePer: number = 54): Generator<string> {
    return this.applyUint8Array(utf8Encoder.encode(data), cutLinePer);
  }

  *applyUint8Array(data: Uint8Array, cutLinePer: number = 54): Generator<string> {
    const bytes = data;
    // 72 / 4 * 3 = 54
    let x = 0;
    let nextLfAt = cutLinePer;
    let line = "";
    const eol = (1 <= cutLinePer) ? "\r\n" : "";
    while (true) {
      const rest = bytes.length - x;
      if (3 <= rest) {
        const word = (bytes[x] << 16) | (bytes[x + 1] << 8) | bytes[x + 2];
        line += base64Table[(word >> 18) & 0x3F] + base64Table[(word >> 12) & 0x3F] + base64Table[(word >> 6) & 0x3F] + base64Table[word & 0x3F];
      }
      else if (2 === rest) {
        const word = (bytes[x] << 16) | (bytes[x + 1] << 8);
        line += base64Table[(word >> 18) & 0x3F] + base64Table[(word >> 12) & 0x3F] + base64Table[(word >> 6) & 0x3F] + "=";
      }
      else if (1 === rest) {
        const word = (bytes[x] << 16);
        line += base64Table[(word >> 18) & 0x3F] + base64Table[(word >> 12) & 0x3F] + "==";
      }
      else {
        break;
      }
      x += 3;
      if (eol && nextLfAt <= x) {
        yield line + eol;
        line = "";
        nextLfAt += cutLinePer;
      }
    }
    if (line) {
      yield line + eol;
    }
  }
}

export class EncodeWordInBase64 {
  encode(word: string): string {
    const base64 = new Base64TransferEncoding();
    const encoded = base64.applyStringAsUtf8(word, 0);
    return `=?utf-8?B?${Array.from(encoded).join("")}?=`;
  }

  encodeIfNeeded(word: string): string {
    if (/^[ !#\$%&@:\-\/\\_A-Za-z0-9]*$/.test(word)) {
      return word;
    }
    else {
      return this.encode(word);
    }
  }
}

export class EmlWriter {
  private _encodeWord: (word: string) => string;
  contentId(cid: string): EmlWriter {
    return this.writeHeader("Content-ID", `<${this.encodeValue(cid)}>`);
  }

  writeContent(factory: Generator<string, any, any>) {
    for (const chunk of factory) {
      this._emitter(chunk);
    }
  }

  writeChunk(chunk: string) {
    this._emitter(chunk);
  }

  writeChunks(chunks: string[]) {
    for (const chunk of chunks) {
      this._emitter(chunk);
    }
  }

  endBoundary(boundary: string): EmlWriter {
    this._emitter(`--${boundary}--\r\n`);
    return this;
  }

  beginBoundary(boundary: string): EmlWriter {
    this._emitter(`--${boundary}\r\n`);
    return this;
  }

  newLine(): EmlWriter {
    this._emitter(`\r\n`);
    return this;
  }

  contentTransferEncoding(type: string): EmlWriter {
    return this.writeHeader("Content-Transfer-Encoding", type);
  }

  mimeVersion1(): EmlWriter {
    return this.writeHeader("MIME-Version", "1.0");
  }

  contentTypeMultipartMixed(boundary: string): EmlWriter {
    return this.writeHeader("Content-Type", `multipart/mixed; boundary="${this.encodeValue(boundary)}"`);
  }

  contentType(type: string): EmlWriter {
    return this.writeHeader("Content-Type", type);
  }

  messageId(messageId: string | undefined): EmlWriter {
    if (messageId) {
      return this.writeHeader("Message-ID", this.encodeValue(messageId));
    }
    else {
      return this;
    }
  }

  date(date: string): EmlWriter {
    return this.writeHeader("Date", this.encodeValue(date));
  }

  subject(subject: string): EmlWriter {
    return this.writeHeader("Subject", this.encodeValue(subject));
  }

  bcc(data: AddressSpec[]): EmlWriter {
    return this.writePersons("Bcc", data);
  }

  cc(data: AddressSpec[]): EmlWriter {
    return this.writePersons("Cc", data);
  }

  to(data: AddressSpec[]): EmlWriter {
    return this.writePersons("To", data);
  }

  private writePersons(headerName: string, persons: AddressSpec[]): EmlWriter {
    if (persons.length === 0) {
      return this;
    }
    else {
      const writer = this.beginHeader(headerName);
      for (const person of persons) {
        writer.value(this.writePerson(person));
      }
      writer.end();
      return this;
    }
  }

  private writePerson(person: AddressSpec): string {
    if (person.name && person.email) {
      return `${this.encodeValue(person.name)} <${person.email}>`;
    }
    else if (person.name) {
      return this.encodeValue(person.name);
    }
    else if (person.email) {
      return `<${this.encodeValue(person.email)}>`;
    }
    else {
      return "";
    }
  }

  from(data: AddressSpec): EmlWriter {
    return this.writeHeader("From", this.writePerson(data));
  }

  private _emitter: ((chunk: string) => void);

  constructor(emitter: (chunk: string) => void, encodeWord: (word: string) => string) {
    this._emitter = emitter;
    this._encodeWord = encodeWord;
  }

  writeHeader(name: string, value: string): EmlWriter {
    this._emitter(`${name}: ${value}\r\n`);
    return this;
  }

  private beginHeader(name: string): HeaderWriter {
    this._emitter(`${name}: `);
    let index = 0;
    const writer: HeaderWriter = {
      value: (value: string) => {
        if (index !== 0) {
          this._emitter(`, `);
        }
        this._emitter(`${value}`);
        index++;
        return writer;
      },
      end: () => {
        this._emitter(`\r\n`);
      }
    };
    return writer;
  }

  private encodeValue(value: string): string {
    return this._encodeWord(value);
  }
}
