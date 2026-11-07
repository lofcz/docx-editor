/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A peer can name one ID in any number of paragraph texts. Indexing them, and asking how each
// shows the ID, must cost about one step per text, not one step per pair of texts.
import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { bodyStoryRoot, readOoxmlPart, WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import { linearizeParagraph } from '../document/paragraph-linear.ts';
import { encodeAttributes } from '../document/paragraph-text.ts';
import { asLogicalId } from '../document/identity.ts';
import { DEFAULT_DOCUMENT_LIMITS } from '../document/limits.ts';
import { InlineIndex } from '../document/paragraph-inline-index.ts';

describe('the inline index under many texts that name one ID', () => {
  test('builds and answers in time near linear in the texts', () => {
    const holders = 20_000;
    const doc = new Y.Doc();
    const nodes = doc.getMap<Y.Map<unknown>>('nodes');
    doc.transact(() => {
      for (let at = 0; at < holders; at += 1) {
        const record = new Y.Map<unknown>();
        nodes.set(`p${at}`, record);
        const text = new Y.Text();
        record.set('inline', text);
        text.insert(0, 'a');
        text.insertEmbed(1, { n: 'X' });
      }
    });
    const order = (left: string, right: string): number =>
      left < right ? -1 : left > right ? 1 : 0;
    const index = new InlineIndex(nodes, DEFAULT_DOCUMENT_LIMITS, () => false, order);
    const started = performance.now();
    expect(index.holdersOf('X').size).toBe(holders);
    let tagged = 0;
    for (let at = 0; at < holders; at += 1) {
      if (index.shownId(asLogicalId(`p${at}`), 'X') !== 'X') tagged += 1;
    }
    const elapsed = performance.now() - started;
    // The first holder by ID shows the ID bare; every other shows it tagged.
    expect(tagged).toBe(holders - 1);
    expect(index.shownId(asLogicalId('p0'), 'X')).toBe('X');
    // A walk over every holder for each one took seconds here.
    expect(elapsed).toBeLessThan(1500);
  });
});

test('only a live record displaces the identity of an inline shell', () => {
  const doc = new Y.Doc();
  try {
    const nodes = doc.getMap<Y.Map<unknown>>('nodes');
    const record = new Y.Map<unknown>();
    nodes.set('control', record);
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:p><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
      {
        name: '/word/document.xml',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      }
    );
    const body = parsed.ok && bodyStoryRoot(parsed.part);
    const paragraphNode = body && body.kind !== 'textValue' && body.children[0];
    if (!paragraphNode || paragraphNode.kind !== 'paragraph')
      throw Error('The paragraph is missing.');
    const control = paragraphNode.children[0]!;
    const authored = { ...paragraphNode, id: 'p', children: [{ ...control, id: 'control' }] };
    const paragraphRecord = new Y.Map<unknown>();
    nodes.set('p', paragraphRecord);
    const text = new Y.Text();
    paragraphRecord.set('inline', text);
    for (const item of linearizeParagraph(authored).items) {
      if (item.kind === 'char')
        text.insert(
          text.length,
          item.value,
          encodeAttributes(item.attributes, 'p', DEFAULT_DOCUMENT_LIMITS)
        );
    }
    let live = true;
    const index = new InlineIndex(
      nodes,
      DEFAULT_DOCUMENT_LIMITS,
      (id) => id === 'control' && !live,
      () => 0
    );
    const paragraph = asLogicalId('p');
    expect(index.holdersOf('control')).toEqual(new Set([paragraph]));
    // A simultaneous record and shell must not have duplicate canonical IDs.
    expect(index.shownId(paragraph, 'control')).not.toBe('control');
    // Undo retains the record for redo, but restores the only live inline owner.
    live = false;
    expect(index.shownId(paragraph, 'control')).toBe('control');
    live = true;
    expect(index.shownId(paragraph, 'control')).not.toBe('control');
  } finally {
    doc.destroy();
  }
});
