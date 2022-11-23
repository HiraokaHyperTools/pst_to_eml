import React, { useCallback, useEffect, useState, } from 'react';
import ReactDOM from 'react-dom';
import { BehaviorSubject } from 'rxjs';
import Container from 'react-bootstrap/Container';
import Button from 'react-bootstrap/Button';
import ButtonGroup from 'react-bootstrap/ButtonGroup';
import Form from 'react-bootstrap/Form';
import ListGroup from 'react-bootstrap/ListGroup';
import Modal from 'react-bootstrap/Modal';
import Nav from 'react-bootstrap/Nav';
import { wrapPstFile, IPFolder, IPItem } from '@hiraokahypertools/pst_to_eml'
import { openPst, PSTFile } from '@hiraokahypertools/pst-extractor'

import 'bootstrap/dist/css/bootstrap.min.css';

let nextNumber = 1;
function nextUniqueKey(): string {
  return `${nextNumber++}`;
}

interface EntryItem {
  key: string;
  display: string;
  messageClass: string;
  entry: IPItem | null;
}

function toItemIsConvertable(source: EntryItem): ItemIsConvertible {
  if (source.messageClass === "IPM.Contact") {
    return {
      canonicalFileName: `${source.display}.vcf`,
      async provider() {
        return source && source.entry && await source.entry.toVCardStr({}) || "";
      },
    };
  }
  else {
    return {
      canonicalFileName: `${source.display}.eml`,
      async provider() {
        return source && source.entry && await source.entry.toEmlStr({}) || "";
      },
    };
  }
}

interface FolderItem {
  key: string;
  display: string;
  entryItemsProvider: () => Promise<EntryItem[]>;
}

interface ItemIsConvertible {
  canonicalFileName: string;
  provider: () => Promise<string>;
}

const folderItemIsNull: FolderItem[] = [{
  key: "null",
  display: "(PST is not loaded yet)",
  async entryItemsProvider() {
    return [];
  },
}];

const folderItemIsEmpty: FolderItem[] = [{
  key: "noOne",
  display: "(Currently no item is available to display)",
  async entryItemsProvider() {
    return [];
  },
}];

const folderItemIsLoading: FolderItem[] = [{
  key: "loading",
  display: "(Loading)",
  async entryItemsProvider() {
    return [];
  },
}];

const entryItemIsNull: EntryItem[] = [{
  key: "null",
  display: "(The folder is not selected)",
  entry: null,
  messageClass: "",
}];

const entryItemIsEmpty: EntryItem[] = [{
  key: "noOne",
  display: "(This folder has no good item to show)",
  entry: null,
  messageClass: "",
}];

const ansiEncodingList = ["utf8", "ascii", "latin1", "armscii8", "big5hkscs", "cp437", "cp737", "cp775", "cp850",
  "cp852", "cp855", "cp856", "cp858", "cp860", "cp861", "cp862", "cp863", "cp864", "cp865", "cp866", "cp869",
  "cp922", "cp932", "cp936", "cp949", "cp950", "cp1046", "cp1124", "cp1125", "cp1129", "cp1133", "cp1161",
  "cp1162", "cp1163", "eucjp", "gb18030", "gbk", "georgianacademy", "georgianps", "hproman8", "iso646cn",
  "iso646jp", "iso88591", "iso88592", "iso88593", "iso88594", "iso88595", "iso88596", "iso88597", "iso88598",
  "iso88599", "iso885910", "iso885911", "iso885913", "iso885914", "iso885915", "iso885916", "koi8r", "koi8ru",
  "koi8t", "koi8u", "maccroatian", "maccyrillic", "macgreek", "maciceland", "macintosh", "macroman",
  "macromania", "macthai", "macturkish", "macukraine", "pt154", "rk1048", "shiftjis", "tcvn", "tis620", "viscii",
  "windows874", "windows1250", "windows1251", "windows1252", "windows1253", "windows1254", "windows1255",
  "windows1256", "windows1257", "windows1258"];

function PSTApp() {
  const fileSubject = new BehaviorSubject<File | null>(null);
  const foldersSubject = new BehaviorSubject<FolderItem[]>(folderItemIsNull);
  const entriesSubject = new BehaviorSubject<EntryItem[]>(entryItemIsNull);
  const previewTextSubject = new BehaviorSubject("");
  const ansiEncodingSubject = new BehaviorSubject("");

  function onChange(input: HTMLInputElement) {
    (input.files?.length === 1) ? fileSubject.next(input.files[0]) : fileSubject.next(null);
  }

  async function openUserPst() {
    entriesSubject.next(entryItemIsEmpty);
    const file = fileSubject.value;
    if (file === null) {
      foldersSubject.next(folderItemIsNull);
      return;
    }
    foldersSubject.next(folderItemIsLoading);
    try {
      const pst = await openPst({
        readFile: async (buffer: ArrayBuffer, offset: number, length: number, position: number) => {
          const blockBlob = file.slice(position, position + length);
          const source = await blockBlob.arrayBuffer();
          new Uint8Array(buffer).set(new Uint8Array(source));
          return blockBlob.size;
        },
        close: async () => {
          fileSubject.next(null);
        },
      }, {
        ansiEncoding: (ansiEncodingSubject.value === "") ? undefined : ansiEncodingSubject.value
      });

      const pstRoot = await wrapPstFile(pst);

      const folderItems: FolderItem[] = [];
      async function walkFolder(folder: IPFolder, prefix: string) {
        folderItems.push({
          key: nextUniqueKey(),
          display: `${prefix} ${await folder.displayName()}`,
          entryItemsProvider: async () => {
            const items: EntryItem[] = [];
            for (let item of await folder.items()) {
              items.push({
                key: nextUniqueKey(),
                display: await item.displayName(),
                messageClass: item.messageClass,
                entry: item,
              });
            }

            return items;
          }
        })
        for (let subFolder of await folder.subFolders()) {
          await walkFolder(subFolder, `${prefix}*`);
        }
      }
      await walkFolder(pstRoot, "");
      foldersSubject.next((folderItems.length !== 0) ? folderItems : folderItemIsEmpty);
    }
    catch (ex) {
      previewTextSubject.next(`${ex}`);
    }
  }

  async function entryOnClick(convertible: ItemIsConvertible) {
    previewTextSubject.next("Loading...");

    try {
      previewTextSubject.next(await convertible.provider());
    } catch (ex) {
      previewTextSubject.next(`${ex}`);
    }
  }

  function FolderSelector() {
    const [folders, setFolders] = useState<FolderItem[]>([]);

    useEffect(
      () => {
        const subscription = foldersSubject.subscribe(
          value => setFolders(value)
        );
        return () => subscription.unsubscribe();
      },
      [foldersSubject]
    );

    async function folderOnChange(index: number) {
      if (index < 0) {
        entriesSubject.next(entryItemIsNull);
        return;
      }
      const hits = await folders[index].entryItemsProvider();
      entriesSubject.next((hits.length !== 0) ? hits : entryItemIsEmpty);
    }

    return (
      <Form.Select onChange={e => folderOnChange(e.target.selectedIndex)}>
        {folders.map(folder => <option key={folder.key}>{folder.display}</option>)}
      </Form.Select>
    );
  }

  function EntriesList() {
    const [entries, setEntries] = useState<EntryItem[]>([]);

    useEffect(
      () => {
        const subscription = entriesSubject.subscribe(
          value => setEntries(value)
        );
        return () => subscription.unsubscribe();
      },
      [entriesSubject]
    );

    return (
      <ListGroup>
        {entries.map(entry =>
          <ListGroup.Item action={entry.entry !== null} key={entry.key} onClick={() => entry.entry && entryOnClick(toItemIsConvertable(entry))}>
            {entry.display}
          </ListGroup.Item>
        )}
      </ListGroup>
    );
  }

  function PreviewModal() {
    const [previewText, setPreviewText] = useState("");

    useEffect(
      () => {
        const subscription = previewTextSubject.subscribe(
          value => setPreviewText(value)
        );
        return () => subscription.unsubscribe();
      }
    );

    return (
      <Modal show={previewText.length !== 0} onHide={() => previewTextSubject.next("")} size='lg'>
        <Modal.Header closeButton>
          <Modal.Title>Preview</Modal.Title>
        </Modal.Header>
        <Modal.Body>
          <Form.Control as="textarea" cols={80} rows={10} value={previewText} readOnly={true} />
        </Modal.Body>
      </Modal>
    );
  }

  function ItemsCount() {
    const [count, setCount] = useState(0);

    useEffect(
      () => {
        const subscription = entriesSubject.subscribe(
          value => {
            if (value === entryItemIsEmpty || value === entryItemIsNull) {
              setCount(0);
            }
            else {
              setCount(value.length);
            }
          }
        );
        return () => subscription.unsubscribe();
      }
    );

    if (count === 0) {
      return <>Items:</>;
    }
    else if (count === 1) {
      return <>1 item:</>;
    }
    else {
      return <>{count} items:</>;
    }
  }

  function Downloader() {
    const [list, setList] = useState<ItemIsConvertible[]>([]);
    const [wip, setWip] = useState(false);

    useEffect(
      () => {
        const subscription = entriesSubject.subscribe(
          value => {
            if (value === entryItemIsEmpty || value === entryItemIsNull) {
              setList([]);
            }
            else {
              setList(value.map(toItemIsConvertable));
            }
          }
        );
        return () => subscription.unsubscribe();
      },
      [entriesSubject]
    );

    function downloadAsFile(fileName: string, text: string) {
      const blob = new Blob([text], { type: "text/plain" });

      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = fileName;
      a.click();
    }

    async function downloadAll() {
      setWip(true);
      try {
        try {
          for (let item of list) {
            downloadAsFile(
              item.canonicalFileName,
              await item.provider()
            );
          }
        }
        catch (ex) {
          previewTextSubject.next(`${ex}`);
        }
      }
      finally {
        setWip(false);
      }
    }

    const count = list.length;

    if (count === 0) {
      return <></>;
    }
    else {
      return (
        <>
          {wip
            ? <Button variant="outline-primary" disabled>In progress...</Button>
            : <Button variant="outline-primary" onClick={() => downloadAll()}>Download {count} items</Button>
          }
        </>
      );
    }
  }

  return <>
    <h1>pst_to_eml demo</h1>
    <Form.Group className="mb-3" controlId='selectPstFile'>
      <Form.Label>Select PST file</Form.Label>
      <Form.Control type="file"
        onChange={e => onChange(e.target as HTMLInputElement)} />
    </Form.Group>
    <Form.Group className="mb-3" controlId='selectAnsiEncoding'>
      <Form.Label>Select ansi encoding</Form.Label>
      <Form.Control placeholder="e.g. windows1251" onChange={e => ansiEncodingSubject.next(e.target.value)} list='ansiEncodingList' />
      <datalist id="ansiEncodingList">
        {ansiEncodingList.map(name => <option key={name} value={name}></option>)}
      </datalist>
    </Form.Group>
    <p><Button onClick={() => openUserPst()}>Open</Button></p>
    <p>Folder:<br /></p>
    <p>
      <FolderSelector />
    </p>
    <p>Folder actions:<br /><Downloader /></p>
    <p><ItemsCount /></p>
    <EntriesList />
    <PreviewModal />
  </>;
}

ReactDOM.render(
  <Container>
    <PSTApp />
  </Container>,
  document.getElementById('root')
);
