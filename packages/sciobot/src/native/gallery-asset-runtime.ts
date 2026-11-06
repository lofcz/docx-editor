import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  contentControlsIn,
  replaceChildren,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { CanonicalEditPlan } from '../diff/plan';
import type { InsertedBlockBinding } from '../diff/inserted-blocks';
import { validGalleryAssetPayload, type GalleryAssetImage } from '../projection/gallery-assets';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { buildNativeGalleryImage } from '../blocks/gallery-native-build';
import { withInlineControlId } from './inline-image-control-id';
import type { ResolvedInlineImage } from './inline-image-resources';

/** Only newly inserted blocks can acquire assets. Existing gallery edits use native commands. */
export function planGalleryAssets(plan: CanonicalEditPlan): readonly InsertedBlockBinding[] {
  const bindings = plan.operations.flatMap((edit) =>
    edit.kind === 'siblings' ? (edit.blocks ?? []) : []
  );
  return bindings.filter((binding) => validGalleryAssetPayload(binding.payload));
}

/** IDs are allocated once for detached validation and the eventual canonical transaction. */
export function galleryAssetCompletion(
  bindings: readonly InsertedBlockBinding[],
  resources: ReadonlyMap<string, ResolvedInlineImage>
): (pkg: OoxmlPackage) => OoxmlPackage {
  const prepared = bindings.map((binding) => ({
    binding,
    images: (binding.payload.attributes.images as GalleryAssetImage[]).map((image) => ({
      image,
      imageId: crypto.randomUUID(),
      inlineId: crypto.randomUUID(),
    })),
  }));
  return (pkg) => {
    let next = pkg;
    const partName = pkg.mainDocumentPart;
    for (const { binding, images } of prepared) {
      let part = next.parts.get(partName)!;
      const matches = contentControlsIn(part.root).filter(
        ({ node }) => contentControlPropertiesOf(node).tag === blockTag(binding.payload.id)
      );
      const content = matches.length === 1 && contentControlContentNodeOf(matches[0]!.node);
      if (!content) throw Error('The gallery asset reservation is missing or ambiguous.');
      const cleared = replaceChildren(part, content.id, []);
      if (!cleared.ok) throw Error('The gallery asset reservation cannot be cleared.');
      next = withPart(next, cleared.part);
      for (const { image, imageId, inlineId } of images) {
        const resource = resources.get(image.asset);
        if (!resource) throw Error('The gallery asset has no prepared resource.');
        const built = buildNativeGalleryImage(next, partName, {
          galleryId: binding.payload.id,
          imageId,
          inlineId,
          reference: image.asset,
          resource: {
            ...resource,
            widthPoints: image.width * 0.75,
            heightPoints: image.height * 0.75,
          },
          caption: image.caption,
          alt: image.alt,
          title: image.title,
        });
        part = built.pkg.parts.get(partName)!;
        const current = contentControlContentNodeOf(
          contentControlsIn(part.root).find(({ node }) => node.id === matches[0]!.node.id)!.node
        )!;
        const changed = replaceChildren(part, current.id, [...current.children, built.control]);
        if (!changed.ok) throw Error('The prepared gallery image cannot be inserted.');
        next = withInlineControlId(
          withPart(built.pkg, changed.part),
          partName,
          built.inlineControlId
        );
      }
      const {
        resourceVersion: _version,
        images: _images,
        ...attributes
      } = binding.payload.attributes;
      const stored = withPayloadRecord(
        next,
        { ...binding.payload, version: 3, attributes },
        binding.label
      );
      if (!stored.ok) throw Error(`Gallery asset ownership refused: ${stored.reason}.`);
      next = stored.pkg;
    }
    return next;
  };
}
