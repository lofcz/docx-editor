import { expect, test } from 'bun:test';
import { blankDocumentBytes } from '@docx-editor.dev/core';
import {
  readOoxmlPackage,
  writeOoxmlPackage,
  withPart,
  paragraphTextOf,
  canonicalOoxmlFingerprint,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { bodyPart } from '../../test/support/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  fixture,
  readBlocks,
  type DocumentBlock,
} from '../../../../examples/react-blocks/src/document';
import { blockTag, withPayloadRecord } from './payload-opc-codec';
import { writeBlockPayload } from './payload-part';
import { buildBlockIndex } from './block-index';

const paragraph = '<w:p><w:r><w:t>Identical content</w:t></w:r></w:p>';
const control = (tag: string, content = paragraph) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
function makePackage(body: string, ids: string[]) {
  const opened = readOoxmlPackage(blankDocumentBytes());
  if (!opened.ok) throw new Error(opened.reason);
  let pkg = withPart(opened.package, bodyPart(body));
  for (const id of ids) {
    const written = withPayloadRecord(
      pkg,
      { id, type: 'lesson_activity', version: 1, attributes: { name: 'Identical', duration: 10 } },
      'Identical'
    );
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return pkg;
}

test('identical-looking blocks retain distinct stable IDs through save and reopen', () => {
  const pkg = makePackage(control(blockTag('first')) + control(blockTag('second')), [
    'first',
    'second',
  ]);
  const fingerprint = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const index = buildBlockIndex(pkg);
  expect([...index.byId.keys()]).toEqual(['first', 'second']);
  expect(index.byId.get('first')!.controlId).not.toBe(index.byId.get('second')!.controlId);
  const opened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!opened.ok) throw new Error(opened.reason);
  const reopened = buildBlockIndex(opened.package);
  expect([...reopened.byId.keys()]).toEqual(['first', 'second']);
  expect(reopened.diagnostics).toEqual([]);
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toBe(fingerprint);
});

test('nested controls resolve their closest registered parent', () => {
  const pkg = makePackage(control(blockTag('parent'), paragraph + control(blockTag('child'))), [
    'parent',
    'child',
  ]);
  const index = buildBlockIndex(pkg);
  expect(index.byId.get('parent')!.parentId).toBeNull();
  expect(index.byId.get('child')!.parentId).toBe('parent');
});

test('duplicates, missing payloads, and orphans leave unrelated blocks unchanged', () => {
  const pkg = makePackage(
    control(blockTag('duplicate')) +
      control(blockTag('duplicate')) +
      control(blockTag('missing')) +
      control(blockTag('good')),
    ['duplicate', 'good', 'orphan']
  );
  const before = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const index = buildBlockIndex(pkg);
  expect([...index.byId.keys()]).toEqual(['good']);
  expect(index.diagnostics.map((item) => [item.code, item.id])).toEqual([
    ['duplicate-reference', 'duplicate'],
    ['missing-payload', 'missing'],
    ['orphan-payload', 'orphan'],
  ]);
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toBe(before);
});

test('malformed-reference diagnostics stay bounded', () => {
  const pkg = makePackage(Array.from({ length: 10 }, () => control('scio:1:')).join(''), []);
  const index = buildBlockIndex(pkg, 3);
  expect(index.diagnostics).toHaveLength(3);
  expect(index.omittedDiagnostics).toBe(7);
  expect(index.byId.size).toBe(0);
});

test('live indexes follow two sessions through edits, undo, deletion, reconnect, and reopen', async () => {
  const pair = await createTwoEditorFixture({
    document: fixture(),
    documentId: 'block-index-replicas',
  });
  const first = pair.editors[0]!;
  const second = pair.editors[1]!;
  const flatten = (blocks: DocumentBlock[]): DocumentBlock[] =>
    blocks.flatMap((block) => [block, ...flatten(block.children)]);
  const block = flatten(readBlocks(first.surface!.session.part())).find(
    (item) => item.type === 'activity'
  )!;
  const label = 'Indexed activity — 15 minutes';
  const old = paragraphTextOf(first.surface!.session.part(), block.headerId)!;
  const index = (pkg: OoxmlPackage) => buildBlockIndex(pkg);
  const current = (peer: number) => index(pair.editors[peer]!.surface!.session.currentPackage());
  try {
    pair.pause();
    expect(
      writeBlockPayload(first, {
        controlId: block.id,
        label,
        payload: {
          id: 'stable',
          type: 'lesson_activity',
          version: 1,
          attributes: { name: 'Indexed activity', duration: 15 },
        },
        fallbackOps: [
          { op: 'deleteText', paragraphId: block.headerId, start: 0, end: old.length },
          { op: 'insertText', paragraphId: block.headerId, offset: 0, text: label },
        ],
      }).success
    ).toBe(true);
    const target = [...block.paragraphs].find((id) =>
      paragraphTextOf(second.surface!.session.part(), id)?.includes('Plan an experiment.')
    )!;
    second.surface!.setSelection({
      anchor: { paragraphId: target, offset: 0 },
      head: { paragraphId: target, offset: 0 },
    });
    second.surface!.insertPlainText('Peer input. ');
    pair.reconnect();
    await pair.settle();
    expect(current(0).byId.get('stable')!.controlId).toBe(current(1).byId.get('stable')!.controlId);
    expect(pair.rooms[0]!.session.remoteSelections().length).toBeGreaterThan(0);
    expect(first.exec({ type: 'undo' }).ok).toBe(true);
    pair.flush();
    expect(current(1).byId.size).toBe(0);
    expect(paragraphTextOf(second.surface!.session.part(), target)).toContain('Peer input.');
    expect(first.exec({ type: 'redo' }).ok).toBe(true);
    pair.flush();
    expect(current(1).byId.get('stable')!.controlId).toBe(block.id);
    pair.pause();
    second.surface!.insertPlainText('Still editing. ');
    expect(
      first.surface!.applyAutomationOps(() => [
        { op: 'removeContentControl', controlId: block.id, keepContent: false },
      ]).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    expect(current(0).byId.size).toBe(0);
    expect(current(1).byId.size).toBe(0);
    expect(current(1).diagnostics).toContainEqual({ code: 'orphan-payload', id: 'stable' });
    const digests = await pair.savedDigests();
    expect(digests[0]).toEqual(digests[1]);
  } finally {
    pair.destroy();
  }
}, 30000);
