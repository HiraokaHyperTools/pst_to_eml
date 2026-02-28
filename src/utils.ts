import * as path from 'node:path';

/**
 * 
 * @internal
 */
export function applyFallbackRecipients(array: Array<any>, fallback: any): Array<any> {
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
export function convertToUint8Array(attachmentStream: ArrayBuffer | undefined): Uint8Array {
  if (attachmentStream instanceof ArrayBuffer) {
    return new Uint8Array(attachmentStream);
  }
  return new Uint8Array(0);
}
