import React, { useCallback, useEffect, useState, } from 'react';
import ReactDOM from 'react-dom';
import { BehaviorSubject } from 'rxjs';
import Container from 'react-bootstrap/Container';
import Button from 'react-bootstrap/Button';
import Form from 'react-bootstrap/Form';
import ListGroup from 'react-bootstrap/ListGroup';
import ListGroupItem from 'react-bootstrap/ListGroupItem';
import Badge from 'react-bootstrap/Badge';
import Modal from 'react-bootstrap/Modal';
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
  entry: IPItem;
}

interface FolderItem {
  key: string;
  display: string;
  entryItemsProvider: () => Promise<EntryItem[]>;
}

function PSTApp() {
  const [file, setFile] = useState<File | null>(null);
  const [folders, setFolders] = useState<FolderItem[]>([]);
  const [entries, setEntries] = useState<EntryItem[]>([]);
  const [showPreview, setShowPreview] = useState(false);
  const [textPreview, setTextPreview] = useState("");

  function onChange(input: HTMLInputElement) {
    (input.files?.length === 1) ? setFile(input.files[0]) : setFile(null);
  }

  async function folderOnChange(index: number) {
    if (index < 0) {
      setEntries([]);
      return;
    }
    setEntries(await folders[index].entryItemsProvider());
  }

  async function openUserPst() {
    setEntries([]);
    if (file === null) {
      setFolders([]);
      return;
    }
    try {
      const pst = await openPst({
        readFile: async (buffer: ArrayBuffer, offset: number, length: number, position: number) => {
          const blockBlob = file.slice(position, position + length);
          const source = await blockBlob.arrayBuffer();
          new Uint8Array(buffer).set(new Uint8Array(source));
          return blockBlob.size;
        },
        close: async () => {
          setFile(null);
        },
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
      setFolders(folderItems);
    }
    catch (ex) {
      setShowPreview(true);
      setTextPreview(`${ex}`);
    }
  }

  async function entryOnClick(item: IPItem) {
    setShowPreview(true);
    setTextPreview("Loading...");

    try {
      if (item.messageClass === "IPM.Contact") {
        setTextPreview(await item.toVCardStr({}));
      }
      else {
        setTextPreview(await item.toEmlStr({}));
      }
    } catch (ex) {
      setTextPreview(`${ex}`);
    }
  }

  return <>
    <h1>pst_to_eml demo</h1>
    <p>Select PST file</p>
    <p>
      <Form.Control
        type="file"
        onChange={e => onChange(e.target as HTMLInputElement)} />
    </p>
    <p><Button onClick={() => openUserPst()}>Open</Button></p>
    <p>Folder:<br /></p>
    <p>
      <Form.Select onChange={e => folderOnChange(e.target.selectedIndex)}>
        {folders.map(folder => <option key={folder.key}>{folder.display}</option>)}
      </Form.Select>
    </p>
    <p>Items:</p>
    <ListGroup>
      {entries.map(entry =>
        <ListGroup.Item action key={entry.key} onClick={() => entryOnClick(entry.entry)}>
          {entry.display}
        </ListGroup.Item>
      )}
    </ListGroup>
    <Modal show={showPreview} onHide={() => setShowPreview(false)} size='lg'>
      <Modal.Header closeButton>
        <Modal.Title>Preview</Modal.Title>
      </Modal.Header>
      <Modal.Body>
        <Form.Control as="textarea" cols={80} rows={10} value={textPreview} readOnly={true} />
      </Modal.Body>
    </Modal>
  </>;
}

ReactDOM.render(
  <Container>
    <PSTApp />
  </Container>,
  document.getElementById('root')
);
