import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import {
  preparedInlineImageResources,
  type PreparedInlineImages,
} from '../native/inline-image-resources';
import { replaceManagedImageResource } from '../native/replace-image-resource';

/** Preserve plan-bound preparation while sharing native replacement with image controls. */
export function completeImageReplacement(
  pkg: OoxmlPackage,
  edit: Extract<CanonicalEdit, { kind: 'inline-image' }>,
  plan: CanonicalEditPlan,
  handle?: PreparedInlineImages
): OoxmlPackage {
  const resource =
    edit.resource.reference !== edit.before.attributes.reference
      ? preparedInlineImageResources(handle, plan)?.get(edit.resource.reference)
      : undefined;
  return replaceManagedImageResource(pkg, edit, resource);
}
