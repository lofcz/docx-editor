import {
  applyTreeOp,
  validateTreeOp,
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  createNodeIdAllocator,
  findNode,
  parentNodeOf,
  paragraphOffsetIndex,
  paragraphTextOf,
  replaceNode,
  withPart,
  authoredProperties,
  AUTHORABLE_RUN_PROPERTIES,
  propertyContainer,
  type OoxmlNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { blockTag } from '../blocks/payload-opc-codec';
import type { InlineTargetMark } from '../diff/inline-target';
import {
  inlineCodeRunChanges,
  inlineCodeCharacterStyles,
  hyperlinkCharacterStyle,
  withInlineCodeStyle,
} from './inline-code';

/** Replace only our owned placeholder, preserving its run, control, and surrounding wrappers. */
export function materializeImageReservation(
  pkg: OoxmlPackage,
  partName: string,
  controlId: string,
  expectedTag: string,
  payloadId: string,
  drawing: Extract<OoxmlNode, { kind: 'drawing' }>,
  marks: readonly InlineTargetMark[]
): OoxmlPackage {
  let next = pkg;
  const part = pkg.parts.get(partName)!,
    control = findNode(part, controlId);
  if (control?.kind !== 'contentControl' || contentControlPropertiesOf(control).tag !== expectedTag)
    throw Error('The image reservation identity changed.');
  const content = contentControlContentNodeOf(control),
    run = content?.children[0];
  if (content?.children.length !== 1 || run?.kind !== 'run')
    throw Error('The image reservation must own exactly one native run.');
  const texts = run.children.filter((node) => node.kind === 'text');
  const text = texts[0];
  if (
    texts.length !== 1 ||
    text?.kind !== 'text' ||
    text.children.length !== 1 ||
    text.children[0]?.kind !== 'textValue' ||
    text.children[0].value !== '\uFFFC' ||
    run.children.some((node) => node.kind !== 'text' && node.kind !== 'runProperties')
  )
    throw Error('The image reservation content changed before materialization.');
  let parent = parentNodeOf(part, controlId);
  for (let depth = 0; parent && parent.kind !== 'paragraph' && depth < 64; depth++)
    parent = parentNodeOf(part, parent.id);
  if (parent?.kind !== 'paragraph')
    throw Error('The image reservation has no containing paragraph.');
  const span = paragraphOffsetIndex(parent).spanOf(control);
  if (
    !span ||
    span.end !== span.start + 1 ||
    paragraphTextOf(part, parent.id)?.slice(span.start, span.end) !== '\uFFFC'
  )
    throw Error('The native image reservation range changed.');
  const refusal = validateTreeOp(part, {
    op: 'deleteText',
    paragraphId: parent.id,
    start: span.start,
    end: span.end,
  });
  if (refusal) throw Error(`Image reservation content is not writable: ${refusal}.`);
  const mint = createNodeIdAllocator(part);
  const fresh = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? { ...node, id: mint() }
      : ({ ...node, id: mint(), children: node.children.map(fresh) } as OoxmlNode);
  const replaced = replaceNode(part, text.id, fresh(drawing));
  if (!replaced.ok)
    throw Error(`Native image replacement refused: ${JSON.stringify(replaced.issues)}.`);
  next = withPart(next, replaced.part);
  const apply = (op: TreeDocOp) => {
    const result = applyTreeOp(next.parts.get(partName)!, op);
    if (!result.ok) throw Error(`Native image formatting refused: ${result.reason}.`);
    next = withPart(next, result.part);
  };
  const inherited = authoredProperties(
    propertyContainer(run, 'runProperties', 'rPr'),
    AUTHORABLE_RUN_PROPERTIES
  ).filter((property) => !['b', 'i', 'strike'].includes(property.localName));
  apply({
    op: 'setRunProperties',
    paragraphId: parent.id,
    start: span.start,
    end: span.end,
    properties: [
      ...inherited,
      ...(['bold', 'italic', 'strike'] as const).map((mark) => ({
        localName: mark === 'bold' ? 'b' : mark === 'italic' ? 'i' : 'strike',
        attributes: { val: marks.includes(mark) ? '1' : '0' },
      })),
    ],
  });
  const code = inlineCodeRunChanges(
    next.parts.get(partName)!,
    { paragraphId: parent.id, start: span.start, end: span.end, code: marks.includes('code') },
    { codeStyles: inlineCodeCharacterStyles(next), hyperlinkStyle: hyperlinkCharacterStyle(next) }
  );
  if (!code.success) throw Error(code.error);
  for (const change of code.changes) {
    if (change.after !== null && marks.includes('code')) {
      const styled = withInlineCodeStyle(next, [...inlineCodeCharacterStyles(next), change.after]);
      if (!styled) throw Error('The image cannot gain its native code style.');
      next = styled;
    }
    apply({
      op: 'setRunProperties',
      paragraphId: parent.id,
      start: change.start,
      end: change.end,
      properties: change.properties,
      characterStyle: change.after,
      ...(change.targetRunIds ? { targetRunIds: change.targetRunIds } : {}),
    });
  }
  apply({
    op: 'setContentControlProperties',
    controlId,
    tag: blockTag(payloadId),
    alias: 'Image',
    lock: 'sdtContentLocked',
  });
  return next;
}
