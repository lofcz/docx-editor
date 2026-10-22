import { expect, test } from 'bun:test';
import { readOoxmlPart, WML_NAMESPACE_URI } from '../package/ooxml-tree.ts';
import { findNode, nodeIndexTestRecorder } from '../package/ooxml-edit.ts';
import { TreeDocumentStore, collectRevisionSites } from '../index.ts';
test('tracking inline ownership keeps the current part index warm', () => {
  const xml = `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${Array.from({ length: 100 }, () => '<w:p><w:r><w:t>word</w:t></w:r></w:p>').join('')}</w:body></w:document>`;
  const read = readOoxmlPart(xml, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!read.ok) throw Error(read.reason);
  const part = read.part,
    store = new TreeDocumentStore(part);
  const body = part.root.children[0];
  if (body.kind === 'textValue') throw Error('body');
  const ids = body.children.map((p) => p.id);
  findNode(part, ids[0]);
  const recorder = nodeIndexTestRecorder();
  recorder.reset();
  const result = store.transact((ctx) => {
    for (const paragraphId of ids)
      ctx.apply({
        op: 'deleteText',
        paragraphId,
        start: 0,
        end: 4,
        revision: { author: 'Bot', date: '2026-10-21T21:00:00Z' },
      });
  });
  expect(result.ok).toBe(true);
  expect(recorder.completeBuilds).toBeLessThanOrEqual(1);
  expect(collectRevisionSites(store.part)).toHaveLength(100);
});
