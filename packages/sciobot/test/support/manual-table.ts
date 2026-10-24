import { docxFixture } from './document';
import type { createTwoEditorFixture } from './collaboration';
import { parentNodeOf, paragraphTextOf } from '@docx-editor.dev/core/store';
import { blockAst, blockToken, blockContext, decideBlocks } from './block-execution';
import { createEditPlan } from '../../src/diff/plan';
import type { TargetAstNode } from '../../src/projection/parse-target';
import { trackSiblingPlan } from '../../src/review/track-blocks';

/** Three real Word rows with imported header behavior, geometry, and extension metadata. */
export function manualTableDocument() {
  const cell = (text: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  return docxFixture(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      ['Header', 'Middle', 'Last']
        .map(
          (label, index) =>
            `<w:tr xmlns:keep="urn:keep" keep:row="${label}"><w:trPr>${index === 0 ? '<w:tblHeader/>' : ''}<w:cantSplit/><w:trHeight w:val="400" w:hRule="atLeast"/></w:trPr>${cell(`${label} A`)}${cell(`${label} B`)}</w:tr>`
        )
        .join('') +
      '</w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
  );
}

/** Recover committed peer text from a rejected AI paragraph inside the first native cell. */
export async function recoverFirstTableCellSource(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>
) {
  const [alice, bob] = pair.editors,
    a = alice!.surface!,
    b = bob!.surface!;
  const part = a.session.part(),
    source = a.session.paragraphIds()[0]!;
  const cell = parentNodeOf(part, source);
  if (cell?.kind !== 'tableCell')
    throw new Error('The first paragraph is not a native cell paragraph.');
  const before = new Set(a.session.paragraphIds());
  const plan = createEditPlan(blockToken(alice!), [
    {
      kind: 'siblings',
      partName: part.name,
      parentId: cell.id,
      beforeId: source,
      afterId: null,
      removedIds: [],
      inserted: blockAst('AI').children as readonly TargetAstNode[],
    },
  ]);
  if (!plan.success) throw new Error(plan.error);
  const staged = trackSiblingPlan(alice!, plan.plan, blockContext);
  if (!staged.success) throw new Error(staged.error);
  await pair.settle();
  const added = b.session.paragraphIds().filter((id) => !before.has(id));
  if (added.length !== 1) throw new Error('The AI paragraph birth is ambiguous.');
  const id = added[0]!;
  pair.pause();
  decideBlocks(alice!, staged.proposals, 'rejectRevision');
  const peer = b.applyAutomationOps(
    () => [{ op: 'insertText', paragraphId: id, offset: 2, text: 'Peer', bias: 'right' }],
    { kind: 'body' },
    undefined,
    false,
    true
  );
  if (peer.rejected || !peer.committed) throw new Error('The concurrent peer edit did not commit.');
  pair.reconnect();
  await pair.settle();
  for (const surface of [a, b])
    if (paragraphTextOf(surface.session.part(), id) !== 'Peer')
      throw new Error('The committed peer text did not recover.');
  return id;
}
