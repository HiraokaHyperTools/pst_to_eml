const { openPstFile } = require('@hiraokahypertools/pst-extractor');
const program = require('commander');

const fs = require('fs');
const path = require('path');
const { wrapPstFile } = require('./lib');

function safety(name) {
  return name.replace(/[\"/\\\\\\?<>\\*:\\|]/g, "_");
}

program
  .command('tree <pstFilePath>')
  .description('Print items tree inside pst file')
  .option('--ansi-encoding <encoding>', 'Set ANSI encoding (used by iconv-lite) for non Unicode text in msg file')
  .action(async (pstFilePath, options) => {
    try {
      const pstFile = await openPstFile(pstFilePath);
      const pst = await wrapPstFile(pstFile);
      try {
        async function walk(node, depth) {
          const prefixTo = (s) => "".padStart(depth, '>') + " " + s;
          for (let item of (await node.items())) {
            console.log(prefixTo("(M) " + await item.displayName()));
          }
          for (let subFolder of (await node.subFolders())) {
            console.log(prefixTo("(f) " + await subFolder.displayName()));
            await walk(subFolder, depth + 1);
          }
        }

        await walk(pst, 0);
      }
      finally {
        pst.close();
      }
    } catch (ex) {
      process.exitCode = 1;
      console.error(ex);
    }
  });

program
  .command('export <pstFilePath> <saveToDir>')
  .description('export items inside pst file')
  .option('--ansi-encoding <encoding>', 'Set ANSI encoding (used by iconv-lite) for non Unicode text in msg file')
  .action(async (pstFilePath, saveToDir, options) => {
    try {
      const pstFile = await openPstFile(pstFilePath, { ansiEncoding: options.ansiEncoding, });
      const pst = await wrapPstFile(pstFile);
      try {
        function safeJoin(one, two) {
          return path.join(one, two);
        }

        async function walk(node, depth, saveToDir) {
          await fs.promises.mkdir(saveToDir, { recursive: true });

          const prefixTo = (s) => "".padStart(depth, '>') + " " + s;
          for (let item of (await node.items())) {
            const displayName = await item.displayName();
            console.log(prefixTo("(M) " + displayName));
            if (item.messageClass === "IPM.Note") {
              const emlStr = await item.toEmlStr({});
              await fs.promises.writeFile(safeJoin(saveToDir, safety(displayName) + ".eml"), emlStr);
            }
            else if (item.messageClass === "IPM.Contact") {
              const str = await item.toVCardStr({});
              await fs.promises.writeFile(safeJoin(saveToDir, safety(displayName) + ".vcf"), str);
            }
            else {
              console.log(`Unprocessed messageClass: ${item.messageClass}`);
            }
          }
          for (let subFolder of (await node.subFolders())) {
            const displayName = await subFolder.displayName();
            console.log(prefixTo("(f) " + displayName));
            await walk(subFolder, depth + 1, safeJoin(saveToDir, safety(displayName)));
          }
        }

        await walk(pst, 0, path.normalize(saveToDir));
      }
      finally {
        pst.close();
      }
    } catch (ex) {
      process.exitCode = 1;
      console.error(ex);
    }
  });

function createUndup() {
  const map = new Map();
  return {
    next: (name) => {
      if (!name) {
        name = "Unnamed";
      }
      for (let x = 0; ; x++) {
        const candidate = x === 0 ? name : `${name} (${x})`;
        if (!map.has(candidate)) {
          map.set(candidate, true);
          return candidate;
        }
      }
    }
  }
}

program
  .command('export-mbox <pstFilePath> <saveToDir>')
  .description('export items inside pst file to mbox format')
  .option('--ansi-encoding <encoding>', 'Set ANSI encoding (used by iconv-lite) for non Unicode text in msg file')
  .action(async (pstFilePath, saveToDir, options) => {
    try {
      const pstFile = await openPstFile(pstFilePath, { ansiEncoding: options.ansiEncoding, });
      const pst = await wrapPstFile(pstFile);
      try {
        function safeJoin(one, two) {
          return path.join(one, two);
        }

        async function walk(node, depth, saveToDir) {
          console.log("walk", saveToDir);
          await fs.promises.mkdir(saveToDir, { recursive: true });

          const undup = createUndup();

          for (let subFolder of (await node.subFolders())) {
            const displayName = undup.next(await subFolder.displayName());
            {
              const mbox = await fs.promises.open(safeJoin(saveToDir, safety(displayName)), 'w')
              try {
                for (let item of (await subFolder.items())) {
                  if (item.messageClass === "IPM.Note" || item.messageClass.indexOf("IPM.Document.") === 0) {
                    const emlStr = await item.toEmlStr({});
                    await mbox.write(`From - _${((new Date()).getTime())}\r\nX-Mozilla-Status: 0001\r\nX-Mozilla-Status2: 00000000\r\nX-Mozilla-Keys:                                                                                 \r\n`);
                    await mbox.write(emlStr);
                    await mbox.write("\r\n");
                  }
                }
              }
              finally {
                await mbox.close();
              }
            }
            await walk(subFolder, depth + 1, safeJoin(saveToDir, safety(displayName) + ".sbd"));
          }
        }

        await walk(pst, 0, path.normalize(saveToDir));
      }
      finally {
        pst.close();
      }
    } catch (ex) {
      process.exitCode = 1;
      console.error(ex);
    }
  });

program
  .parse(process.argv);
