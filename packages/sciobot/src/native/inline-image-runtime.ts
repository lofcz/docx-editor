import { materializeImageReservation } from './inline-image-reservation';
import { embedInlineImage } from './inline-image-media';
import { inlinePicture } from './inline-picture';
import {
  contentControlsIn,
  contentControlPropertiesOf,
  escapeXml,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createInlineImagePayload } from '../blocks/inline-image-codec';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
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
        for (const { atom, id } of bindings) {
          const part = next.parts.get(partName)!;
          const controls = contentControlsIn(part.root).filter(
            ({ node }) => contentControlPropertiesOf(node).tag === PREFIX + id
          );
          if (controls.length !== 1)
            throw Error('The image reservation must resolve to exactly one native control.');
          const control = controls[0]!.node;
          const store = readPayloadStore(next);
          if (store.status === 'invalid' || (store.status === 'ready' && store.entries.has(id)))
            throw Error('The image payload identity is unavailable.');
          const resource = resources.get(atom.resource.reference)!;
          const embedded = embedInlineImage(next, partName, id, resource);
          next = embedded.pkg;
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
          next = materializeImageReservation(
            next,
            partName,
            control.id,
            PREFIX + id,
            id,
            drawing,
            atom.marks
          );
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
          const stored = withPayloadRecord(next, payload, 'Image');
          if (!stored.ok) throw Error(`Image provenance storage refused: ${stored.reason}.`);
          next = stored.pkg;
        }
        return next;
      },
    },
  };
}
