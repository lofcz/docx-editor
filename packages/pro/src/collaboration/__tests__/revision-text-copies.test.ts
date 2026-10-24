/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { readOoxmlPackage, type OoxmlNode, type OoxmlElement } from '@docx-editor.dev/core/store';
import { zipDocument } from './document-peer-support';
import { linearizeParagraph } from '../document/paragraph-linear.ts';
import { encodeAttributes } from '../document/paragraph-text.ts';
import { tokensOfText } from '../document/paragraph-text-diff.ts';
import { textIdentities } from '../document/paragraph-text-identity.ts';
import { pairedRevisionSources, revisionCopyPositions } from '../document/revision-text-copies.ts';
import { DEFAULT_DOCUMENT_LIMITS as limits } from '../document/limits.ts';

function firstParagraph(node: OoxmlNode): OoxmlElement | undefined {
  if (node.kind === 'textValue') return undefined;
  if (node.kind === 'paragraph') return node;
  for (const child of node.children) {
    const found = firstParagraph(child);
    if (found) return found;
  }
}

function fixture(author = 'Alice', date = '2026-09-01T00:00:00Z') {
  const loaded = readOoxmlPackage(
    zipDocument(`<w:p>
    <w:del w:id="1" w:author="Alice" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>x</w:delText></w:r></w:del>
    <w:ins w:id="2" w:author="${author}" w:date="${date}"><w:r><w:t>x</w:t></w:r></w:ins>
  </w:p>`)
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const root = loaded.package.parts.get(loaded.package.mainDocumentPart)!.root;
  const paragraph = firstParagraph(root);
  if (!paragraph) throw new Error('paragraph required');
  const doc = new Y.Doc();
  const text = doc.getText('text');
  for (const item of linearizeParagraph(paragraph).items) {
    if (item.kind === 'char')
      text.insert(text.length, item.value, encodeAttributes(item.attributes, paragraph.id, limits));
  }
  return { doc, text, id: paragraph.id };
}

test('native source identity, rather than equal text, establishes a review pair', () => {
  const { doc, text, id } = fixture();
  try {
    const after = tokensOfText(text);
    expect(pairedRevisionSources([{ id, after }], limits).size).toBe(0);
    const copied = after.map((token) => ({ ...token, source: '41:100' }));
    expect([...pairedRevisionSources([{ id, after: copied }], limits)]).toEqual(['41:100']);
    expect(pairedRevisionSources([{ id, after: [...copied, copied[0]!] }], limits).size).toBe(0);
    const identities = textIdentities(text);
    expect(revisionCopyPositions(text, identities, limits, id).size).toBe(0);
    const pairs = revisionCopyPositions(
      text,
      { ...identities, ids: ['41:100', '41:100'] },
      limits,
      id
    );
    expect(pairs.get(0)?.role.kind).toBe('del');
    expect(pairs.get(0)?.other).toBe(1);
    expect(pairs.get(1)?.role.kind).toBe('ins');
  } finally {
    doc.destroy();
  }
});

for (const [author, date] of [
  ['Bob', '2026-09-01T00:00:00Z'],
  ['Alice', '2026-09-02T00:00:00Z'],
]) {
  test(`different review ownership refuses pairing: ${author}/${date}`, () => {
    const { doc, text, id } = fixture(author, date);
    try {
      const after = tokensOfText(text).map((token) => ({ ...token, source: '41:100' }));
      expect(pairedRevisionSources([{ id, after }], limits).size).toBe(0);
      const identities = textIdentities(text);
      expect(
        revisionCopyPositions(text, { ...identities, ids: ['41:100', '41:100'] }, limits, id).size
      ).toBe(0);
    } finally {
      doc.destroy();
    }
  });
}
