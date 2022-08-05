const assert = require('assert');
const { openPstFile } = require('@hiraokahypertools/pst-extractor');
const { wrapPstFile } = require('../lib/index');

const fs = require('fs');
const path = require('path');

const baseDir = __dirname;

describe("msgInMsgInMsg.pst", function () {
  it("items() multiple items", async function () {
    const pstFilePath = path.join(baseDir, "msgInMsgInMsg.pst");
    const pstFile = await openPstFile(pstFilePath);
    const pst = await wrapPstFile(pstFile);

    const folders1 = await pst.subFolders();
    const folder2 = folders1[0];
    const folders3 = await folder2.subFolders();
    const folder4 = folders3[0];
    const items5a = await folder4.items();
    const items5b = await folder4.items();
    const items5c = await folder4.items();
    assert.strictEqual(items5a.length, 1);
    assert.strictEqual(items5b.length, 1);
    assert.strictEqual(items5c.length, 1);
    pstFile.close();
  });
});
describe("tree traversal tests", function () {
  async function walk(folder, stat, onItem) {
    stat.structure += "[";
    for (let subFolder of await folder.subFolders()) {
      for (let item of await subFolder.items()) {
        stat.structure += "E(";
        stat.structure += `${item.messageClass},`;
        const emlStr = await onItem(item);
        stat.structure += `${emlStr && emlStr.length || "NA"},`;
        stat.structure += ")";
      }

      await walk(subFolder, stat, onItem);
    }
    stat.structure += "]";
  }
  async function traversalTest(pstPath, onItem) {
    const stat = { structure: "", };
    const pstFilePath = path.join(baseDir, pstPath);
    const pstFile = await openPstFile(pstFilePath, { ansiEncoding: "cp932" });
    const pst = await wrapPstFile(pstFile);
    await walk(pst, stat, onItem);
    await pstFile.close();
    return stat;
  }

  async function itemPrinter(item) {
    if (item.messageClass === "IPM.Note") {
      const emlStr = await item.toEmlStr({});
      assert.notEqual(
        emlStr,
        undefined
      );
      assert.notEqual(
        await item.toEmlBuffer({}),
        undefined
      );
      return emlStr;
    }
    else if (item.messageClass === "IPM.Contact") {
      const str = await item.toVCardStr({});
      assert.notEqual(
        str,
        undefined
      );
      //console.log(str);
      //await require('fs').promises.writeFile(`${new Date().getTime()}.vcf`, str);
      return str;
    }
    else {
      throw new Error(`Unknown messageClass: ${item.messageClass}`);
    }
  }

  it("msgInMsgInMsg.pst", async function () {
    const stat = await traversalTest("msgInMsgInMsg.pst", itemPrinter);
    assert.deepEqual(stat, { structure: "[[E(IPM.Note,3053,)[[]][]]]" });
  });
  it("contacts.pst", async function () {
    const stat = await traversalTest("contacts.pst", itemPrinter);
    assert.deepEqual(stat, { structure: "[[[[[]]E(IPM.Contact,504,)[]][][]]]" });
  });
  it("contacts97-2002.pst", async function () {
    const stat = await traversalTest("contacts97-2002.pst", itemPrinter);
    assert.deepEqual(stat, { structure: "[[[[]E(IPM.Contact,504,)[]][][]]]" });
  });
});
