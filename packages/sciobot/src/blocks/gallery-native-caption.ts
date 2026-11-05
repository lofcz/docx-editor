import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  findNode,
  readOoxmlPart,
  w14RootPrefix,
  WML_NAMESPACE_URI,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { withFreshIds } from './fresh-ids';
import { readNativeGalleryImage } from './gallery-native-record';
import type { GalleryImageState } from './gallery';
import { reviewGroupHistoryRefusal } from '../review/group-history';

/** Seed one shared caption owner before the grouped image can receive concurrent first edits. */
export function emptyNativeGalleryCaption(
  part: OoxmlPart,
  imageId: string,
  allocate: () => string,
  used: Set<string>
) {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr></w:p></w:body></w:document>`,
    { name: '/sciobot-native-gallery-caption.xml', contentType: part.contentType }
  );
  const body = parsed.ok && bodyStoryRoot(parsed.part);
  const paragraph = body && body.kind !== 'textValue' && body.children[0];
  if (!paragraph || paragraph.kind !== 'paragraph')
    throw Error('The native caption template is invalid.');
  return withFreshIds(paragraph, allocate, {
    prefix: w14RootPrefix(part.root),
    used,
    seed: `gallery-caption:${imageId}`,
  });
}

/** Change only the differing native text span. Native character ownership remains authoritative. */
export function updateNativeGalleryCaption(
  editor: DocxEditorInstance,
  image: GalleryImageState,
  caption: string
) {
  const surface = editor.surface!;
  const history = reviewGroupHistoryRefusal(surface.collaborationSession());
  if (history) throw Error(history);
  const part = surface.session.part(),
    pkg = surface.session.currentPackage();
  const control = findNode(part, image.controlId)!;
  const native = readNativeGalleryImage(pkg, part, control);
  if (!native?.caption)
    throw Error(
      'The native caption owner is missing. Ungroup and group these images before editing captions.'
    );
  const previous = String(native.attributes.caption),
    next = caption.replace(/\r\n|\r/g, '\n');
  if (previous === next)
    return { success: true as const, revision: surface.session.packageRevision() };
  const before = [...previous],
    after = [...next];
  let prefix = 0,
    suffix = 0;
  while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix])
    prefix++;
  while (
    suffix < before.length - prefix &&
    suffix < after.length - prefix &&
    before[before.length - suffix - 1] === after[after.length - suffix - 1]
  )
    suffix++;
  const start = before.slice(0, prefix).join('').length;
  const end = previous.length - before.slice(before.length - suffix).join('').length;
  const text = after
    .slice(prefix, after.length - suffix)
    .join('')
    .replace(/\n/g, '\v');
  const ops: TreeDocOp[] = [];
  if (end > start) ops.push({ op: 'deleteText', paragraphId: native.caption.id, start, end });
  if (text) ops.push({ op: 'insertText', paragraphId: native.caption.id, offset: start, text });
  const result = surface.applyAutomationOps(() => ops, { kind: 'body' }, undefined, false, true);
  if (result.rejected || !result.committed)
    throw Error(String(result.reason ?? 'The native caption did not change.'));
  return { success: true as const, revision: surface.session.packageRevision() };
}
