import { expect, test } from 'bun:test';
import { bodyParagraphs, allParagraphs, treeToDoc, docToTreeOps } from '../tree-binding.ts';
import { part, apply } from '../../store/__tests__/tracked-edit-fixture.ts';
import { runWithTransactionActor } from '../../store/package/actor-scoped-ids.ts';
import { findNode } from '../../store/index.ts';

const p = '<w:p><w:r><w:t>Kept</w:t></w:r></w:p>';
const pending =
  '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="AI"/></w:rPr></w:pPr><w:ins w:id="1" w:author="AI"><w:r><w:t>New</w:t></w:r></w:ins></w:p>';

test('body and all-story projection caches exclude an empty retired sibling without deleting its identity', () => {
  const source = part(p + pending);
  const [kept, old] = allParagraphs(source).map((node) => node.id);
  const next = runWithTransactionActor('alice', () => apply(source, { op: 'rejectAllRevisions' }));
  expect(findNode(next, old!)).not.toBeNull();
  for (let n = 0; n < 2; n++) {
    expect(bodyParagraphs(next).map((node) => node.id)).toEqual([kept!]);
    expect(allParagraphs(next).map((node) => node.id)).toEqual([kept!]);
    const doc = treeToDoc(next);
    expect(doc.childCount).toBe(1);
    expect(doc.textContent).toBe('Kept');
    expect(docToTreeOps(next, doc)).toMatchObject({ ok: true, ops: [] });
  }
  const peer = apply(next, { op: 'insertText', paragraphId: old!, offset: 0, text: 'Peer' });
  expect(allParagraphs(peer).map((node) => node.id)).toEqual([kept!, old!]);
  expect(treeToDoc(peer).textContent).toBe('KeptPeer');
  expect(allParagraphs(next).map((node) => node.id)).toEqual([kept!]);
});
