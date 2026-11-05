import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  findNode,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  serializeOoxmlPart,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { codeDocument, openCode } from '../../test/support/code';
import { buildBlockIndex } from './block-index';
import { nativeCodeSource } from './code-native-source';
import { codeSourceOps } from './code-source-ops';

const SOURCE = 'one\nmiddle\nlast';
const apply = (part: OoxmlPart, op: TreeDocOp) => {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(`${op.op}: ${result.reason}`);
  return result.part;
};

for (const value of [
  'one Alice\nmiddle\nlast',
  'onemiddle\nlast',
  'one\n\tmid\ndle\nlast',
  'one\nnew\nmiddle\nlast',
  'onemid\nnew\ndle\nlast',
  '\n\tone\n\nmiddle\nlast\n',
  '',
  'replacement\n\tvalue\n',
])
  test(`canonical code revisions preserve both source views and decisions: ${JSON.stringify(value)}`, async () => {
    const pkg = openCode(await codeDocument());
    const block = buildBlockIndex(pkg).byId.get('code')!;
    let part = apply(pkg.parts.get(block.partName)!, {
      op: 'setContentControlProperties',
      controlId: block.controlId,
      lock: 'unlocked',
    });
    const original = serializeOoxmlPart(part);
    const source = nativeCodeSource(part, findNode(part, block.controlId)!)!;
    const revision = { author: 'Alice', date: '2026-11-05T12:00:00.000Z' };
    const operations = codeSourceOps(source, value, revision);
    const before = part;
    const scopes: { paragraphId: string; scope: string }[] = [];
    for (const paragraph of source.paragraphs) {
      if (
        !operations.some(
          (op) =>
            'paragraphId' in op &&
            op.paragraphId === paragraph.id &&
            ['insertText', 'insertTab', 'deleteText', 'splitParagraphMany'].includes(op.op)
        )
      )
        continue;
      const prepared = prepareTrackedRunOrigins(part, paragraph.id, [
        { start: 0, end: source.lines[source.paragraphs.indexOf(paragraph)]!.length },
      ]);
      part = prepared.part;
      if (prepared.scope) scopes.push({ paragraphId: paragraph.id, scope: prepared.scope });
    }
    for (const op of operations) part = apply(part, op);
    for (const { paragraphId, scope } of scopes)
      part = sealTrackedRunOrigins(part, paragraphId, scope, before);
    for (const paragraph of nativeCodeSource(part, findNode(part, block.controlId)!, 'committed')!
      .paragraphs)
      if (!source.paragraphs.some((original) => original.id === paragraph.id))
        for (const { scope } of scopes)
          part = sealTrackedRunOrigins(part, paragraph.id, scope, before);
    expect(nativeCodeSource(part, findNode(part, block.controlId)!, 'committed')?.value).toBe(
      SOURCE
    );
    expect(nativeCodeSource(part, findNode(part, block.controlId)!, 'working')?.value).toBe(value);
    expect(nativeCodeSource(part, findNode(part, block.controlId)!)).toBeNull();
    const accepted = apply(part, { op: 'acceptAllRevisions' });
    const rejected = apply(part, { op: 'rejectAllRevisions' });
    expect(nativeCodeSource(accepted, findNode(accepted, block.controlId)!)?.value).toBe(value);
    expect(nativeCodeSource(rejected, findNode(rejected, block.controlId)!)?.value).toBe(SOURCE);
    expect(serializeOoxmlPart(rejected)).toBe(original);
  });
