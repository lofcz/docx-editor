import { activeMdxRegistry } from '../registry/active';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  cleanupOrphanImageMedia,
  contentControlPropertiesOf,
  contentControlsIn,
  findNode,
  isStoryPart,
  parentNodeOf,
  replaceNode,
  resolveImageRelationship,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { blockIdFromTag, readPayloadStore } from '../blocks/payload-opc-codec';
import type { InlineRemoval } from '../diff/inline-removal';
import { removableInlinePayload } from '../diff/inline-removal';
import { jsonDataKey } from '../diff/target-key';
import { refuseWrite } from '../diff/fragment-seeds';

/** Validate an immutable atom source without applying preview operations to an observed transaction. */
export function validateInlineRemovalSource(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  edit: InlineRemoval
) {
  const block = buildBlockIndex(pkg).byId.get(edit.blockId),
    control = findNode(part, edit.controlId);
  if (
    !block ||
    block.partName !== part.name ||
    block.controlId !== edit.controlId ||
    !control ||
    control.kind !== 'contentControl' ||
    !removableInlinePayload(block.payload) ||
    jsonDataKey(block.payload) !== jsonDataKey(edit.before) ||
    canonicalOoxmlFingerprint(control) !== edit.fingerprint
  )
    refuseWrite('The removed atom no longer matches its native source.', 'stale-revision');
  let parent = parentNodeOf(part, control!.id);
  while (parent && parent.kind !== 'paragraph') {
    // Only ordinary native link wrappers may surround an inline atom.
    if (
      parent.kind !== 'hyperlink' &&
      !(parent.kind === 'generic' && parent.localName === 'sdtContent')
    )
      refuseWrite('The removed atom has an unsupported native owner.', 'unsupported-operation');
    parent = parentNodeOf(part, parent.id);
  }
  if (parent?.id !== edit.paragraphId)
    refuseWrite('The removed atom belongs to another paragraph.', 'stale-revision');
  let references = 0;
  for (const story of pkg.parts.values())
    if (isStoryPart(story))
      for (const { node } of contentControlsIn(story.root))
        if (blockIdFromTag(contentControlPropertiesOf(node).tag ?? '') === edit.blockId)
          references++;
  if (references !== 1)
    refuseWrite('The removed atom has ambiguous native ownership.', 'ambiguous-target');
  const children = (node: OoxmlNode): OoxmlNode[] => [
    node,
    ...(node.kind === 'textValue' ? [] : node.children.flatMap(children)),
  ];
  if (contentControlsIn(control!).some(({ node }) => node.id !== edit.controlId))
    refuseWrite('The removed atom contains another native control.', 'unsupported-operation');
  const relationships = children(control!).flatMap((node) =>
    node.kind === 'textValue'
      ? []
      : node.attributes
          .filter(
            (a) =>
              a.namespaceUri ===
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships' &&
              a.localName === 'embed'
          )
          .map((a) => a.value)
  );
  const media =
    activeMdxRegistry().components.get(edit.before.type)?.mediaPart?.(edit.before) ?? null;
  if (typeof media === 'string') {
    const resolved =
      relationships.length === 1 &&
      resolveImageRelationship(pkg.relationships.get(part.name), part.name, relationships[0]!);
    if (!resolved || resolved.mode !== 'internal' || resolved.partName !== media)
      refuseWrite('The removed image resource no longer matches its provenance.', 'stale-revision');
  }
  return { control: control!, relationships, media };
}

/** Prepare native deletion without flattening a locked fallback into text. */
export function lowerInlineRemoval(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  edit: InlineRemoval
): {
  readonly ops: readonly TreeDocOp[];
  readonly part: OoxmlPart;
  readonly complete: (pkg: OoxmlPackage) => OoxmlPackage;
} {
  const { control, relationships, media } = validateInlineRemovalSource(pkg, part, edit);
  const ops: TreeDocOp[] = [];
  let preview = part;
  const apply = (op: TreeDocOp) => {
    const result = applyTreeOp(preview, op);
    if (!result.ok) refuseWrite(`Native atom removal refused: ${result.reason}.`, 'apply-failed');
    preview = result.part;
    ops.push(op);
  };
  if (contentControlPropertiesOf(control!).lock === 'sdtContentLocked')
    apply({ op: 'setContentControlProperties', controlId: edit.controlId, lock: 'unlocked' });
  apply({ op: 'removeContentControl', controlId: edit.controlId, keepContent: false });
  return {
    ops,
    part: preview,
    complete(working) {
      const store = readPayloadStore(working);
      if (
        store.status !== 'ready' ||
        store.entries.get(edit.blockId)?.decoded.status !== 'supported' ||
        jsonDataKey(store.entries.get(edit.blockId)!.decoded.value) !== jsonDataKey(edit.before)
      )
        refuseWrite('The removed atom payload record disappeared.', 'stale-revision');
      const payloadPart = working.parts.get(store.part.partName)!;
      const record = payloadPart.root.children.find(
        (node) =>
          node.kind !== 'textValue' &&
          node.attributes.some(
            (a) =>
              a.namespaceUri === '' && a.localName === 'id' && a.value === `scio_${edit.blockId}`
          )
      );
      if (!record || record.kind === 'textValue')
        refuseWrite('The removed atom payload record is unavailable.', 'apply-failed');
      // Keep an empty identity shell so concurrent history can restore the same record.
      const retired = replaceNode(payloadPart, record.id, {
        ...record,
        attributes: [],
        children: [],
      });
      if (!retired.ok) refuseWrite('The removed atom payload could not retire.', 'apply-failed');
      let next = withPart(working, retired.part);
      if (typeof media === 'string')
        for (const relationship of relationships)
          next = cleanupOrphanImageMedia(next, part.name, media, relationship);
      return next;
    },
  };
}
