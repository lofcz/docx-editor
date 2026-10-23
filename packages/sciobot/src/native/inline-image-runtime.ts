import { inlinePicture } from './inline-picture';
import {
  applyTreeOp,
  contentControlsIn,
  contentControlPropertiesOf,
  parentNodeOf,
  paragraphOffsetIndex,
  paragraphTextOf,
  withEmbeddedImage,
  withPart,
  escapeXml,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createInlineImagePayload } from '../blocks/inline-image-codec';
import { blockTag, readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { NativeEditFailure } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import type { InsertedInlineAtom } from '../diff/inserted-atoms';
import {
  planInlineDrawings,
  preparedInlineImageResources,
  type PreparedInlineImages,
} from './inline-image-resources';
type Drawing = Extract<InsertedInlineAtom, { kind: 'drawing' }>;
export interface InlineImageRuntime {
  render(atom: InsertedInlineAtom): { readonly xml: string; readonly length: number } | null;
  reservation(atom: Drawing, paragraphId: string): TreeDocOp;
  complete(pkg: OoxmlPackage): OoxmlPackage;
}
const PREFIX = 'scio:image-reservation:';
export function createInlineImageRuntime(
  plan: CanonicalEditPlan,
  handle?: PreparedInlineImages
): { readonly success: true; readonly runtime?: InlineImageRuntime } | NativeEditFailure {
  const drawings = planInlineDrawings(plan);
  if (!drawings.length) return { success: true };
  const resources = preparedInlineImageResources(handle, plan);
  if (!resources)
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'Image resources require revision-bound executor preparation.',
    };
  if (drawings.some((atom) => atom.hyperlink || atom.marks.length))
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'Marked and linked images need explicit native ownership.',
    };
  const bindings = drawings.map((atom) => ({ atom, id: crypto.randomUUID() }));
  const byResource = new WeakMap(
    bindings.map((binding) => [binding.atom.resource, binding] as const)
  );
  const bindingOf = (atom: Drawing) => {
    const binding = byResource.get(atom.resource);
    if (!binding || !resources.has(atom.resource.reference))
      throw Error('The image reservation has no prepared resource.');
    return binding;
  };
  return {
    success: true,
    runtime: {
      render(atom) {
        if (atom.kind !== 'drawing') return null;
        const binding = bindingOf(atom);
        return {
          xml: `<w:sdt><w:sdtPr><w:tag w:val="${escapeXml(PREFIX + binding.id)}"/></w:sdtPr><w:sdtContent><w:r><w:t>\uFFFC</w:t></w:r></w:sdtContent></w:sdt>`,
          length: 1,
        };
      },
      reservation(atom, paragraphId) {
        return {
          op: 'insertInlineContentControl',
          paragraphId,
          offset: atom.offset,
          tag: PREFIX + bindingOf(atom).id,
          text: '\uFFFC',
          alias: 'Image',
        };
      },
      complete(pkg) {
        let next = pkg;
        const partName = pkg.mainDocumentPart;
        const apply = (op: TreeDocOp) => {
          const result = applyTreeOp(next.parts.get(partName)!, op);
          if (!result.ok) throw Error(`Native image operation refused: ${result.reason}.`);
          next = withPart(next, (result as { part: OoxmlPart }).part);
        };
        for (const { atom, id } of bindings) {
          let part = next.parts.get(partName)!;
          const controls = contentControlsIn(part.root).filter(
            ({ node }) => contentControlPropertiesOf(node).tag === PREFIX + id
          );
          if (controls.length !== 1)
            throw Error('The image reservation must resolve to exactly one native control.');
          const control = controls[0]!.node,
            paragraph = parentNodeOf(part, control.id);
          if (paragraph?.kind !== 'paragraph')
            throw Error('The native image reservation must belong to a paragraph.');
          const span = paragraphOffsetIndex(paragraph).spanOf(control);
          if (
            !span ||
            span.end !== span.start + 1 ||
            paragraphTextOf(part, paragraph.id)?.slice(span.start, span.end) !== '\uFFFC'
          )
            throw Error('The native image reservation changed before materialization.');
          const store = readPayloadStore(next);
          if (store.status === 'invalid' || (store.status === 'ready' && store.entries.has(id)))
            throw Error('The image payload identity is unavailable.');
          const resource = resources.get(atom.resource.reference)!;
          const embedded = withEmbeddedImage(next, partName, resource);
          if (!embedded.ok) throw Error(`Image resource embedding refused: ${embedded.reason}.`);
          next = embedded.pkg;
          apply({ op: 'removeContentControl', controlId: control.id, keepContent: true });
          const drawing = inlinePicture({
            docPrId: embedded.docPrId,
            relationshipId: embedded.relationshipId,
            widthPoints: resource.widthPoints,
            heightPoints: resource.heightPoints,
            title: atom.resource.title,
            alt: atom.resource.alt,
          });
          if (drawing.kind !== 'drawing')
            throw Error('The native image builder returned no drawing.');
          apply({ op: 'insertDrawing', paragraphId: paragraph.id, offset: span.start, drawing });
          apply({
            op: 'deleteText',
            paragraphId: paragraph.id,
            start: span.start + 1,
            end: span.start + 2,
          });
          apply({
            op: 'insertContentControl',
            paragraphId: paragraph.id,
            start: span.start,
            end: span.start + 1,
            type: 'richText',
            tag: blockTag(id),
            alias: atom.resource.alt || 'Image',
            lock: 'sdtContentLocked',
          });
          const payload = createInlineImagePayload(id, {
            reference: atom.resource.reference,
            alt: atom.resource.alt,
            title: atom.resource.title,
            contentHash: resource.contentHash,
            mediaPartName: embedded.partName.startsWith('/')
              ? embedded.partName
              : '/' + embedded.partName,
            widthPoints: resource.widthPoints,
            heightPoints: resource.heightPoints,
          });
          const stored = withPayloadRecord(next, payload, atom.resource.alt || 'Image');
          if (!stored.ok) throw Error(`Image provenance storage refused: ${stored.reason}.`);
          next = stored.pkg;
        }
        return next;
      },
    },
  };
}
