import { openPstFile } from '@hiraokahypertools/pst-extractor';
import { IPFolder, MsgConverterOptions, PFolder, PItem, PRoot, wrapPstFile } from '../src/index';
import fs from 'node:fs';
import path from 'node:path';
import { cwd } from 'node:process';

const baseDir = __dirname;

const generateTestData = false;

interface Result {
  structure: string;
}

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
    expect(items5a.length).toBe(1);
    expect(items5b.length).toBe(1);
    expect(items5c.length).toBe(1);
    pstFile.close();
  });
});
describe("actual extraction tests", function () {
  it("msgInMsgInMsg.pst", async function () {
    await extract("msgInMsgInMsg.pst");
  });
  it("contacts.pst", async function () {
    await extract("contacts.pst");
  });
  it("contacts97-2002.pst", async function () {
    await extract("contacts97-2002.pst");
  });
  it("threeItems.pst", async function () {
    await extract("threeItems.pst");
  });
  it("Hello CJK.pst", async function () {
    await extract("Hello CJK.pst");
  });

  async function extract(pstPath: string) {
    const pstFilePath = path.join(baseDir, pstPath);
    const pstFile = await openPstFile(pstFilePath, { ansiEncoding: "ms932" });
    const pst = await wrapPstFile(pstFile);
    await walk(pst, path.join(cwd(), "tests", "exports", path.basename(pstPath) + ".export"));
    await pstFile.close();
  }

  async function walk(folder: IPFolder, outputDir: string) {
    for (const subFolder of await folder.subFolders()) {
      const subOutputDir = path.join(outputDir, await subFolder.displayName());
      await fs.promises.mkdir(subOutputDir, { recursive: true });
      for (const item of await subFolder.items()) {
        const exportUnits = await exportItems(item);
        for (const { body, suffix } of exportUnits) {
          const filePath = path.join(subOutputDir, `${item.primaryNodeId}${suffix}`);
          await applyResult(filePath, body);
        }
      }

      await walk(subFolder, subOutputDir);
    }
  }

  const exportOpt: MsgConverterOptions = {
    baseBoundary: "a38c05f5-fa4a-4c6c-afc5-29da7f34bb5f",
    altBoundary: "9b9ae776-d66d-4e54-a476-1310e57090fd",
  };
  const exportOptNested: MsgConverterOptions = {
    baseBoundary: "a38c05f5-fa4a-4c6c-afc5-29da7f34bb5f",
    altBoundary: "9b9ae776-d66d-4e54-a476-1310e57090fd",
    allowNestedEml: true,
  };

  async function exportItems(item: PItem): Promise<{ body: string, suffix: string }[]> {
    if (item.messageClass === "IPM.Note") {
      const emlStr = await item.toEmlStr(exportOpt);
      const nestedEmlStr = await item.toEmlStr(exportOptNested);
      return [
        { body: emlStr, suffix: ".eml" },
        { body: nestedEmlStr, suffix: ".nested.eml" },
      ];
    }
    else if (item.messageClass === "IPM.Contact") {
      const str = await item.toVCardStr({});
      return [{ body: str, suffix: ".vcf" }];
    }
    else {
      throw new Error(`Unknown messageClass: ${item.messageClass}`);
    }
  }
});
describe("tree traversal tests", function () {
  it("msgInMsgInMsg.pst", async function () {
    const stat = await traversalTest("msgInMsgInMsg.pst");
    expect(stat).toEqual({ structure: "[[E(IPM.Note,2097188,))[[]][]]]" });
  });
  it("contacts.pst", async function () {
    const stat = await traversalTest("contacts.pst");
    expect(stat).toEqual({ structure: "[[[[[]]E(IPM.Contact,2097188,))[]][][]]]" });
  });
  it("contacts97-2002.pst", async function () {
    const stat = await traversalTest("contacts97-2002.pst");
    expect(stat).toEqual({ structure: "[[[[]E(IPM.Contact,2097188,))[]][][]]]", });
  });

  async function walk(folder: IPFolder, stat: Result) {
    stat.structure += "[";
    for (const subFolder of await folder.subFolders()) {
      for (const item of await subFolder.items()) {
        stat.structure += "E(";
        stat.structure += `${item.messageClass},${item.primaryNodeId},)`;
        stat.structure += ")";
      }

      await walk(subFolder, stat);
    }
    stat.structure += "]";
  }
  async function traversalTest(pstPath: string) {
    const stat: Result = { structure: "", };
    const pstFilePath = path.join(baseDir, pstPath);
    const pstFile = await openPstFile(pstFilePath, { ansiEncoding: "ms932" });
    const pst = await wrapPstFile(pstFile);
    await walk(pst, stat);
    await pstFile.close();
    return stat;
  }
});

async function applyResult(filePath: string, body: string) {
  if (generateTestData) {
    await fs.promises.writeFile(filePath, body);
  }
  else {
    const expected = await fs.promises.readFile(filePath, "utf-8");
    expect(body).toBe(expected);
  }
}
