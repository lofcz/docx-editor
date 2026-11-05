// Run with the DOM preload and an immutable engine snapshot that predates gallery version 3.
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { semanticDigest } from '@docx-editor.dev/core/store';
import {
  nativeGalleryDocument,
  nativeGalleryDrawings,
  galleryDecodePort,
} from '../support/native-gallery';
import { groupAdjacentImages } from '../../src/blocks/gallery-group';
import { readPayloadRecords } from '../../src/blocks/payload-opc-codec';
import { openGallery } from '../support/gallery';

const [snapshot, revision, destination] = process.argv.slice(2);
if (!snapshot || !revision || !/^[a-f0-9]{9,40}$/.test(revision) || !destination)
  throw Error('Provide the old snapshot, exact revision, and fixture destination.');
const storePath = resolve(snapshot, 'packages/core/dist/store.js');
const privatePath = resolve(snapshot, 'packages/sciobot/dist/index.js');
const oldStore = await import(pathToFileURL(storePath).href);
const oldPrivate = await import(pathToFileURL(privatePath).href);
const editor = createDocxEditor({
  container: document.createElement('div'),
  document: await nativeGalleryDocument('First\r\nSecond'),
  imageDecodePort: galleryDecodePort,
});
try {
  const grouped = groupAdjacentImages(editor, {
    drawingNodeId: nativeGalleryDrawings(editor.surface!.session.currentPackage())[0]!.drawing.id,
  });
  assert.equal(grouped.success, true);
  const input = new Uint8Array(await editor.save());
  const opened = oldStore.readOoxmlPackage(input);
  assert.equal(opened.ok, true);
  const records = oldPrivate.readPayloadRecords(opened.package);
  assert.equal(records.status, 'ready');
  const opaque = [...records.entries.values()].filter(
    (entry: any) => entry.decoded.status === 'unsupported-version'
  );
  assert.equal(opaque.length, 4);
  for (const record of opaque as any[]) {
    assert.equal(record.decoded.version, 3);
    assert.equal(record.decoded.raw, record.raw);
    const result = oldPrivate.withPayloadRecord(opened.package, record.decoded.value, record.label);
    assert.equal(result.ok, false);
    assert.equal(result.pkg, opened.package);
  }
  const output = new Uint8Array(oldStore.writeOoxmlPackage(opened.package));
  assert.deepEqual(
    semanticDigest(openGallery(input).parts.values()),
    semanticDigest(openGallery(output).parts.values())
  );
  assert.deepEqual(
    [...readRecords(input)].map(([id, entry]) => [id, entry.raw]),
    [...readRecords(output)].map(([id, entry]) => [id, entry.raw])
  );
  await mkdir(destination, { recursive: true });
  await writeFile(resolve(destination, 'native-gallery-v3-input.docx'), input);
  await writeFile(resolve(destination, 'native-gallery-v3-old-output.docx'), output);
  const hash = async (path: string) =>
    createHash('sha256')
      .update(new Uint8Array(await Bun.file(path).arrayBuffer()))
      .digest('hex');
  await writeFile(
    resolve(destination, 'native-gallery-v3-old-writer.json'),
    JSON.stringify(
      {
        revision,
        coreStoreSha256: await hash(storePath),
        privateBundleSha256: await hash(privatePath),
        inputSha256: createHash('sha256').update(input).digest('hex'),
        outputSha256: createHash('sha256').update(output).digest('hex'),
      },
      null,
      2
    ) + '\n'
  );
  console.log(
    `Verified old reader and writer at ${revision}: four opaque records, native image and caption equality.`
  );
} finally {
  editor.destroy();
}
function readRecords(bytes: Uint8Array) {
  const records = readPayloadRecords(openGallery(bytes));
  if (records.status !== 'ready') throw Error('Missing saved payload records');
  return records.entries;
}
