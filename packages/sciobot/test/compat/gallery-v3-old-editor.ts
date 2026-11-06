// Execute a neighboring text edit with an immutable runtime that predates payload version 3.
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  storyParagraphs,
  bodyStoryRoot,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { readPayloadRecords } from '../../src/blocks/payload-opc-codec';
import { openGallery } from '../support/gallery';

const [snapshot, revision, fixtureDirectory] = process.argv.slice(2);
if (!snapshot || !revision || !/^[a-f0-9]{9,40}$/.test(revision) || !fixtureDirectory)
  throw Error('Provide the old snapshot, exact revision, and fixture directory.');
const editorPath = resolve(snapshot, 'packages/core/dist/editor.js');
const storePath = resolve(snapshot, 'packages/core/dist/store.js');
const privatePath = resolve(snapshot, 'packages/sciobot/dist/index.js');
const oldEditor = await import(pathToFileURL(editorPath).href);
const oldStore = await import(pathToFileURL(storePath).href);
const oldPrivate = await import(pathToFileURL(privatePath).href);
const input = new Uint8Array(
  await readFile(resolve(fixtureDirectory, 'native-gallery-v3-input.docx'))
);
const editor = oldEditor.createDocxEditor({
  container: document.createElement('div'),
  document: input,
});
try {
  const surface = editor.surface;
  const pkg = surface.session.currentPackage();
  const records = oldPrivate.readPayloadRecords(pkg);
  assert.equal(records.status, 'ready');
  const opaque = [...records.entries.values()].filter(
    (entry: any) => entry.decoded.status === 'unsupported-version'
  );
  assert.equal(opaque.length, 4);
  const part = pkg.parts.get(pkg.mainDocumentPart);
  const tail = oldStore
    .storyParagraphs(oldStore.bodyStoryRoot(part))
    .find((paragraph: any) => oldStore.paragraphTextOf(part, paragraph.id) === 'Tail');
  assert.ok(tail);
  const edit = surface.applyAutomationOps(
    () => [{ op: 'insertText', paragraphId: tail.id, offset: 4, text: ' edited in Runtime G' }],
    { kind: 'body' },
    undefined,
    false,
    true
  );
  assert.equal(edit.rejected, false);
  assert.equal(edit.committed, true);
  const output = new Uint8Array(await editor.save());
  const before = openGallery(input),
    after = openGallery(output);
  const gallery = [...buildBlockIndex(before).byId.values()].find(
    (block) => block.payload.type === 'image_gallery'
  )!;
  const nextGallery = buildBlockIndex(after).byId.get(gallery.id)!;
  assert.equal(
    canonicalOoxmlFingerprint(findNode(before.parts.get(gallery.partName)!, gallery.controlId)!),
    canonicalOoxmlFingerprint(
      findNode(after.parts.get(nextGallery.partName)!, nextGallery.controlId)!
    )
  );
  const media = (value: typeof before) =>
    [...value.partBytes].filter(([name]) => name.includes('/media/'));
  assert.deepEqual(media(after), media(before));
  const previousRecords = readPayloadRecords(before),
    nextRecords = readPayloadRecords(after);
  assert.equal(previousRecords.status, 'ready');
  assert.equal(nextRecords.status, 'ready');
  if (previousRecords.status !== 'ready' || nextRecords.status !== 'ready')
    throw Error('Missing payload records');
  assert.deepEqual(
    [...previousRecords.entries.values()].map((record) => record.raw),
    [...nextRecords.entries.values()].map((record) => record.raw)
  );
  const afterPart = after.parts.get(after.mainDocumentPart)!;
  assert.equal(
    storyParagraphs(bodyStoryRoot(afterPart)!)
      .map((paragraph) => paragraphTextOf(afterPart, paragraph.id))
      .at(-1),
    'Tail edited in Runtime G'
  );
  const reopened = oldStore.readOoxmlPackage(output);
  assert.equal(reopened.ok, true);
  const reopenedRecords = oldPrivate.readPayloadRecords(reopened.package);
  assert.equal(reopenedRecords.status, 'ready');
  assert.equal(
    [...reopenedRecords.entries.values()].filter(
      (entry: any) => entry.decoded.status === 'unsupported-version'
    ).length,
    4
  );
  await writeFile(resolve(fixtureDirectory, 'native-gallery-v3-old-edited.docx'), output);
  const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
  await writeFile(
    resolve(fixtureDirectory, 'native-gallery-v3-old-editor.json'),
    JSON.stringify(
      {
        revision,
        coreEditorSha256: hash(await readFile(editorPath)),
        coreStoreSha256: hash(await readFile(storePath)),
        privateBundleSha256: hash(await readFile(privatePath)),
        inputSha256: hash(input),
        outputSha256: hash(output),
        operation: {
          op: 'insertText',
          paragraphText: 'Tail',
          offset: 4,
          text: ' edited in Runtime G',
        },
      },
      null,
      2
    ) + '\n'
  );
  console.log(
    `Verified neighboring native text edit with Runtime G ${revision}; opaque gallery records and fallback remain unchanged.`
  );
} finally {
  editor.destroy();
}
