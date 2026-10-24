import type { OoxmlPackage, OoxmlPart } from '@docx-editor.dev/core/store';
import type { CanonicalEdit } from '../diff/plan';
import type { InlineImageRuntime } from '../native/inline-image-runtime';
import { lowerInlineRemoval } from '../native/remove-inline-atom';
import { lowerRetainedAtoms } from './direct-inline-atoms';

export function lowerInlineChange(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  edit: Extract<CanonicalEdit, { kind: 'inline-atoms' | 'inline-remove' }>,
  images?: InlineImageRuntime,
  codeStyles: readonly string[] = []
) {
  return edit.kind === 'inline-remove'
    ? lowerInlineRemoval(pkg, part, edit)
    : lowerRetainedAtoms(pkg, part, edit, images, codeStyles);
}
