import type { EditOptions } from '../package/ooxml-edit.ts';
import type {
  OoxmlHyperlinkNode,
  OoxmlParagraphNode,
  OoxmlPart,
  OoxmlNode,
} from '../package/ooxml-tree.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';
import type { TreeDocOp, TreeOpResult } from './tree-op-types.ts';
export type InlineControlRoot = OoxmlParagraphNode | OoxmlHyperlinkNode;
type Insert = Omit<Extract<TreeDocOp, { op: 'insertInlineContentControl' }>, 'op' | 'paragraphId'>;
type Placement = (
  part: OoxmlPart,
  root: InlineControlRoot,
  op: Insert,
  options?: EditOptions
) => TreeOpResult;
/** An interior tracked insertion keeps the existing hyperlink as its native parent. */
export function trackedControlLinkPlacement(
  part: OoxmlPart,
  paragraph: InlineControlRoot,
  op: Insert,
  options: EditOptions | undefined,
  insert: Placement
): TreeOpResult | null {
  if (!op.revision || paragraph.kind !== 'paragraph') return null;
  const index = paragraphOffsetIndex(paragraph),
    pending: OoxmlNode[] = [...paragraph.children];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'textValue') continue;
    const span = index.spanOf(node);
    if (!span || op.offset <= span.start || op.offset >= span.end) continue;
    if (node.kind === 'hyperlink') {
      const placed = insert(part, node, { ...op, offset: op.offset - span.start }, options);
      return placed.ok
        ? { ...placed, effect: { ...placed.effect, dirty: [paragraph.id] } }
        : placed;
    }
    for (const child of node.children) pending.push(child);
  }
  return null;
}
