import * as path from 'node:path';

/**
 * 
 * @internal
 */
export function formatFrom(senderName: string, senderEmail: string): string {
  if (senderName) {
    return `${senderName} <${senderEmail}>`;
  }
  else {
    return `${senderEmail}`;
  }
}

/**
 * 
 * @internal
 */
export function applyFallbackRecipients(array, fallback) {
  if (array.length === 0) {
    array.push(fallback);
  }
  return array;
}

/**
 * 
 * @internal
 */
export function changeFileExtension(fileName: string, newExt: string): string {
  const parsed = path.parse(fileName);
  return (parsed.dir ? parsed.dir + path.sep : "") + parsed.name + newExt;
}

/**
 * 
 * @internal
 */
export function convertToBuffer(attachmentStream: ArrayBuffer | undefined): Buffer {
  if (attachmentStream instanceof ArrayBuffer) {
    return Buffer.from(attachmentStream);
  }
  return Buffer.alloc(0);
}

/**
 * 
 * @internal
 */
export function formatAddress(name: string, email: string): string | null {
  if (name) {
    return `${email} <${name}>`;
  }
  else {
    return email;
  }
}
