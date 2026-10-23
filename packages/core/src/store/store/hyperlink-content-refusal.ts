import type { OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { hyperlinkPreservedControls } from './hyperlink-control-envelope.ts';
import { rejectContentEdit } from './tree-op-validate-controls.ts';
import type { TreeDocOp } from './tree-op-types.ts';

export function hyperlinkContentRefusal(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  op: Extract<TreeDocOp, { op: 'insertHyperlink' }>
) {
  const preservedWhole = hyperlinkPreservedControls(
    paragraph,
    op,
    op.styleId === undefined && op.revision === undefined
  );
  return rejectContentEdit(part, paragraph, op.start, op.end, { preservedWhole });
}
