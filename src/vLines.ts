
/**
 * 
 * @internal
 */
export function convertVLines(vLines: [string | string[], string | string[]][]): string {
  function toText(vCell: string[] | string): string {
    return Array.isArray(vCell)
      ? vCell.join(';')
      : vCell + ""
  }

  function vEscape(text: string): { more: string, text: string } {
    if (text.match(/\r|\n|\t/)) {
      return {
        more: ";ENCODING=QUOTED-PRINTABLE",
        text: text
          .replace(/\t/g, "=09")
          .replace(/\r/g, "=0D")
          .replace(/\n/g, "=0A")
        ,
      };
    }
    else {
      return { more: "", text };
    }
  }

  const lines = [];
  vLines.forEach(
    vLine => {
      if (vLine[1] === undefined) {
        return;
      }
      const printer = vEscape(toText(vLine[1]));

      lines.push(toText(vLine[0]) + printer.more + ":" + printer.text);
    }
  )
  return lines.join("\n");
}
