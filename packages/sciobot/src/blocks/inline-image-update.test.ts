import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  canonicalOoxmlFingerprint,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  writeOoxmlPackage,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryDocument, openGallery } from '../../test/support/gallery';
import { buildBlockIndex } from './block-index';
import { ungroupGallery } from './gallery-ungroup';
import { readManagedImage, updateInlineImage, type InlineImageUpdate } from './inline-image-update';
import { removeBlock } from './remove';
import { readPayloadRecords, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { nativeImageResource } from './native-image-fields';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { collectBlockClipboardData, pasteWithBlocks } from '../react-blocks/clipboard';
import { lowerInlineRemoval } from '../native/remove-inline-atom';

const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (value) => value.charCodeAt(0)
);
const decodePort: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
async function documentWithImages() {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: galleryDocument(),
    imageDecodePort: decodePort,
  });
  try {
    const result = await ungroupGallery(editor, {
      blockId: 'gallery',
      decodePort,
      resolve: async () => ({ bytes: png, mime: 'image/png' }),
    });
    if (!result.success) throw Error(result.error);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
  }
}
const descendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants)),
];
function imageOf(pkg: OoxmlPackage, id?: string) {
  const block = [...buildBlockIndex(pkg).byId.values()].find(
    (block) => block.payload.type === 'inline_image' && (!id || block.id === id)
  )!;
  const drawing = descendants(findNode(pkg.parts.get(block.partName)!, block.controlId)!).find(
    (node) => node.kind === 'drawing'
  )!;
  return readManagedImage(pkg, block.partName, drawing.id)!;
}
function state(pkg: OoxmlPackage, id?: string) {
  const image = imageOf(pkg, id);
  const nodes = descendants(image.drawing);
  const native = (kind: string, name: string) => {
    const node = nodes.find((node) => node.kind === kind);
    return node?.kind === 'textValue'
      ? undefined
      : node?.attributes.find((attribute) => attribute.localName === name)?.value;
  };
  expect(contentControlPropertiesOf(image.control).lock).toBe('sdtContentLocked');
  expect(Number(native('drawingExtent', 'cx'))).toBe(
    Math.round(image.payload.attributes.widthPoints * 12700)
  );
  expect(Number(native('drawingExtent', 'cy'))).toBe(
    Math.round(image.payload.attributes.heightPoints * 12700)
  );
  expect(native('drawingDocPr', 'descr') ?? '').toBe(image.payload.attributes.alt);
  return image.payload.attributes;
}
function update(editor: DocxEditorInstance, change: InlineImageUpdate['change']) {
  const session = editor.surface!.session;
  const image = imageOf(session.currentPackage());
  return updateInlineImage(editor, {
    drawingNodeId: image.drawing.id,
    expectedRevision: session.packageRevision(),
    change,
  });
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`managed image size and alt remain native through peer edits, history, deletion, and reopen ${clientIds.join('/')}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await documentWithImages(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: decodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const original = state(alice!.surface!.session.currentPackage());
      const originalStore = readPayloadRecords(alice!.surface!.session.currentPackage());
      expect(originalStore.status).toBe('ready');
      pair.pause();
      expect(update(alice!, { type: 'size', widthPoints: 120, heightPoints: 60 }).success).toBe(
        true
      );
      expect(update(bob!, { type: 'alt', alt: 'Peer description' }).success).toBe(true);
      pair.reconnect(0);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(editor.surface!.session.currentPackage())).toMatchObject({
          widthPoints: 120,
          heightPoints: 60,
          alt: 'Peer description',
        });
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(editor.surface!.session.currentPackage())).toMatchObject({
          widthPoints: original.widthPoints,
          heightPoints: original.heightPoints,
          alt: 'Peer description',
        });
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(openGallery(new Uint8Array(await editor.save())))).toMatchObject({
          widthPoints: 120,
          heightPoints: 60,
          alt: 'Peer description',
        });
      pair.pause();
      expect(update(alice!, { type: 'alt', alt: 'Alice description' }).success).toBe(true);
      expect(update(bob!, { type: 'alt', alt: 'Bob description' }).success).toBe(true);
      pair.reconnect(1);
      await pair.settle();
      expect(state(alice!.surface!.session.currentPackage())).toEqual(
        state(bob!.surface!.session.currentPackage())
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(editor.surface!.session.currentPackage()).alt).toBe('Bob description');
      for (let round = 0; round < 3; round++) {
        pair.pause();
        expect(
          update(alice!, { type: 'size', widthPoints: 140 + round, heightPoints: 70 }).success
        ).toBe(true);
        expect(update(alice!, { type: 'alt', alt: `Alice round ${round}` }).success).toBe(true);
        expect(update(bob!, { type: 'alt', alt: `Bob round ${round}` }).success).toBe(true);
        pair.reconnect(round % 2 === 0 ? 0 : 1);
        await pair.settle();
        expect(state(alice!.surface!.session.currentPackage())).toEqual(
          state(bob!.surface!.session.currentPackage())
        );
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(state(editor.surface!.session.currentPackage()).alt).toBe(`Bob round ${round}`);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(state(alice!.surface!.session.currentPackage())).toEqual(
          state(bob!.surface!.session.currentPackage())
        );
      }
      for (const editor of pair.editors) {
        expect(readPayloadRecords(editor.surface!.session.currentPackage())).toEqual(originalStore);
        const revision = readEditorRevision(editor, 'image-ui', 'native')!;
        const view = await readWorkingView(editor.surface!.session, revision);
        if (!view.success) throw Error(view.error);
        expect(view.view.projection.text).toContain(
          state(editor.surface!.session.currentPackage()).alt
        );
      }
      const deleted = imageOf(bob!.surface!.session.currentPackage());
      pair.pause();
      expect(update(alice!, { type: 'size', widthPoints: 200, heightPoints: 100 }).success).toBe(
        true
      );
      expect(removeBlock(bob!, { controlId: deleted.control.id, mode: 'delete' })).toMatchObject({
        success: true,
      });
      pair.reconnect(0);
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          buildBlockIndex(editor.surface!.session.currentPackage()).byId.has(deleted.block.id)
        ).toBe(false);
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(
          buildBlockIndex(editor.surface!.session.currentPackage()).byId.has(deleted.block.id)
        ).toBe(true);
        expect(state(openGallery(new Uint8Array(await editor.save())))).toMatchObject({
          widthPoints: 200,
          heightPoints: 100,
        });
      }
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          buildBlockIndex(editor.surface!.session.currentPackage()).byId.has(deleted.block.id)
        ).toBe(false);
    } finally {
      pair.destroy();
    }
  });

test('managed image edits refuse collaboration without the production ownership scope before writing', async () => {
  const pair = await createTwoEditorFixture({ document: await documentWithImages() });
  try {
    const editor = pair.editors[0]!;
    const image = imageOf(editor.surface!.session.currentPackage());
    const canonical = pair.editors.map((entry) =>
      writeOoxmlPackage(entry.surface!.session.currentPackage())
    );
    const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    expect(update(editor, { type: 'alt', alt: 'No ownership' })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(removeBlock(editor, { controlId: image.control.id, mode: 'delete' })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    await pair.settle();
    expect(
      pair.editors.map((entry) => writeOoxmlPackage(entry.surface!.session.currentPackage()))
    ).toEqual(canonical);
    expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
  } finally {
    pair.destroy();
  }
});

for (const privateData of [true, false])
  test(`image clipboard keeps native metadata after edits and reopen; private=${privateData}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: await documentWithImages(),
      imageDecodePort: decodePort,
    });
    try {
      expect(update(editor, { type: 'alt', alt: 'Copied native description' }).success).toBe(true);
      expect(update(editor, { type: 'size', widthPoints: 90, heightPoints: 45 }).success).toBe(
        true
      );
      const surface = editor.surface!;
      const original = imageOf(surface.session.currentPackage());
      const paragraph = parentNodeOf(surface.session.part(), original.control.id)!;
      const ids = surface.session.paragraphIds();
      const next = ids[ids.indexOf(paragraph.id) + 1]!;
      surface.setSelection({
        anchor: { paragraphId: paragraph.id, offset: 0 },
        head: { paragraphId: next, offset: 0 },
      });
      const flavors = surface.copyFlavours();
      const data = collectBlockClipboardData(editor)!;
      expect(data.blocks).toHaveLength(1);
      expect(data.blocks[0]!.payload.attributes.alt).toBe('Copied native description');
      const tail = ids.at(-1)!;
      const position = {
        paragraphId: tail,
        offset: paragraphTextOf(surface.session.part(), tail)!.length,
      };
      surface.setSelection({ anchor: position, head: position });
      expect(
        pasteWithBlocks(
          editor,
          { ...flavors, data: privateData ? data : null },
          { createId: () => 'copied-image' }
        )
      ).toBe(true);
      expect(state(surface.session.currentPackage(), 'copied-image')).toMatchObject({
        alt: 'Copied native description',
        widthPoints: 90,
        heightPoints: 45,
      });
      const copied = imageOf(surface.session.currentPackage(), 'copied-image');
      expect(
        updateInlineImage(editor, {
          drawingNodeId: copied.drawing.id,
          expectedRevision: surface.session.packageRevision(),
          change: { type: 'alt', alt: 'Independent copy' },
        }).success
      ).toBe(true);
      expect(state(surface.session.currentPackage(), original.block.id).alt).toBe(
        'Copied native description'
      );
      expect(state(openGallery(new Uint8Array(await editor.save())), 'copied-image').alt).toBe(
        'Independent copy'
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(state(surface.session.currentPackage(), 'copied-image').alt).toBe(
        'Copied native description'
      );
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(state(surface.session.currentPackage(), 'copied-image').alt).toBe('Independent copy');
      expect(removeBlock(editor, { controlId: copied.control.id, mode: 'delete' }).success).toBe(
        true
      );
      expect(buildBlockIndex(surface.session.currentPackage()).byId.has('copied-image')).toBe(
        false
      );
      expect(state(surface.session.currentPackage(), original.block.id).alt).toBe(
        'Copied native description'
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(state(openGallery(new Uint8Array(await editor.save())), 'copied-image').alt).toBe(
        'Independent copy'
      );
    } finally {
      editor.destroy();
    }
  });

for (const protection of ['contentLocked', 'sdtContentLocked', 'sdtLocked', 'dataBinding'])
  test(`managed image commands preserve external ${protection} restrictions`, async () => {
    const files = unzipSync(await documentWithImages());
    const property =
      protection === 'dataBinding'
        ? '<w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000001}" w:xpath="/value"/>'
        : `<w:lock w:val="${protection}"/>`;
    files['word/document.xml'] = strToU8(
      strFromU8(files['word/document.xml']!)
        .replace('<w:body>', `<w:body><w:sdt><w:sdtPr>${property}</w:sdtPr><w:sdtContent>`)
        .replace('</w:body>', '</w:sdtContent></w:sdt></w:body>')
    );
    const pair = await createTwoEditorFixture({ document: zipSync(files), ownedAttributes: true });
    try {
      const editor = pair.editors[0]!;
      const before = editor.surface!.session.currentPackage();
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      const image = imageOf(before);
      expect(update(editor, { type: 'alt', alt: 'Description' })).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(update(editor, { type: 'size', widthPoints: 100, heightPoints: 50 }).success).toBe(
        false
      );
      expect(removeBlock(editor, { controlId: image.control.id, mode: 'delete' }).success).toBe(
        false
      );
      await pair.settle();
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
    } finally {
      pair.destroy();
    }
  });

test('generic payload writes cannot introduce a second image field authority', async () => {
  const pkg = openGallery(await documentWithImages());
  const image = imageOf(pkg);
  const updated = withPayloadRecord(
    pkg,
    { ...image.payload, attributes: { ...image.payload.attributes, alt: 'Private overwrite' } },
    'Image'
  );
  expect(updated.ok).toBe(false);
  expect(updated.pkg).toBe(pkg);
  expect(state(pkg).alt).toBe('First image');
});

test('native image removal uses current fields and unchanged provenance', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: await documentWithImages(),
    imageDecodePort: decodePort,
  });
  try {
    expect(update(editor, { type: 'alt', alt: 'Edited removal' }).success).toBe(true);
    expect(update(editor, { type: 'size', widthPoints: 160, heightPoints: 80 }).success).toBe(true);
    const surface = editor.surface!;
    const pkg = surface.session.currentPackage();
    const image = imageOf(pkg);
    const lowered = lowerInlineRemoval(pkg, surface.session.part(), {
      kind: 'inline-remove',
      paragraphId: parentNodeOf(surface.session.part(), image.control.id)!.id,
      controlId: image.control.id,
      blockId: image.block.id,
      before: image.payload,
      fingerprint: canonicalOoxmlFingerprint(image.control),
    });
    expect(
      surface.applyAutomationOps(
        () => lowered.ops,
        { kind: 'body' },
        [lowered.complete],
        false,
        true,
        'after'
      ).committed
    ).toBe(true);
    expect(buildBlockIndex(surface.session.currentPackage()).byId.has(image.block.id)).toBe(false);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(state(openGallery(new Uint8Array(await editor.save())), image.block.id)).toMatchObject({
      alt: 'Edited removal',
      widthPoints: 160,
      heightPoints: 80,
    });
  } finally {
    editor.destroy();
  }
});

for (const invalid of ['external-resource', 'different-resource', 'invalid-geometry'])
  test(`unowned image projections preserve private metadata and refuse edits: ${invalid}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: await documentWithImages(),
      imageDecodePort: decodePort,
    });
    let loaded: DocxEditorInstance | undefined;
    try {
      expect(update(editor, { type: 'alt', alt: 'Native label' }).success).toBe(true);
      expect(update(editor, { type: 'size', widthPoints: 120, heightPoints: 60 }).success).toBe(
        true
      );
      const image = imageOf(editor.surface!.session.currentPackage());
      const resource = nativeImageResource(
        editor.surface!.session.currentPackage(),
        image.part.name,
        image.control
      )!;
      const files = unzipSync(new Uint8Array(await editor.save()));
      if (invalid === 'invalid-geometry') {
        files['word/document.xml'] = strToU8(
          strFromU8(files['word/document.xml']!).replace('cx="1524000"', 'cx="invalid"')
        );
      } else {
        const target =
          invalid === 'external-resource'
            ? 'Target="https://untrusted.example/image.png" TargetMode="External"'
            : 'Target="media/different-unowned-image.png"';
        files['word/_rels/document.xml.rels'] = strToU8(
          strFromU8(files['word/_rels/document.xml.rels']!).replace(
            `Target="${resource.raw}"`,
            target
          )
        );
      }
      loaded = createDocxEditor({
        container: document.createElement('div'),
        document: zipSync(files),
        imageDecodePort: decodePort,
      });
      const session = loaded.surface!.session;
      const before = session.currentPackage();
      const record = readPayloadStore(before);
      if (record.status !== 'ready') throw Error('Missing test metadata.');
      const payload = record.entries.get(image.block.id)!.decoded;
      if (payload.status !== 'supported') throw Error('Missing image payload.');
      expect(payload.value.attributes.alt).toBe('First image');
      expect(payload.value.attributes.widthPoints).not.toBe(120);
      const control = buildBlockIndex(before).byId.get(image.block.id)!;
      const drawing = descendants(findNode(session.part(), control.controlId)!).find(
        (node) => node.kind !== 'textValue' && node.localName === 'drawing'
      )!;
      expect(readManagedImage(before, session.part().name, drawing.id)).toBeNull();
      expect(
        updateInlineImage(loaded, {
          drawingNodeId: drawing.id,
          expectedRevision: session.packageRevision(),
          change: { type: 'alt', alt: 'Refused' },
        })
      ).toMatchObject({ success: false, code: 'invalid-input' });
      expect(session.currentPackage()).toBe(before);
      expect(readPayloadRecords(openGallery(new Uint8Array(await loaded.save())))).toEqual(
        readPayloadRecords(before)
      );
    } finally {
      loaded?.destroy();
      editor.destroy();
    }
  });
