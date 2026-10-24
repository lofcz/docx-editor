import {
  WML_NAMESPACE_URI,
  contentControlPropertiesOf,
  findNode,
  replaceNode,
  withPart,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import { validateInlineRemovalSource } from '../native/remove-inline-atom';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

/** Preserve the complete control and payload until the reviewer decides its native deletion. */
export function stageTrackedInlineRemoval(
  writer: TrackedStageWriter,
  edit: Extract<CanonicalEdit, { kind: 'inline-remove' }>,
  attribution: TrackedStageAttribution
): void {
  const pkg = writer.currentPackage(),
    part = pkg.parts.get(edit.partName);
  if (!part) refuseWrite('The removed atom story is missing.', 'invalid-document');
  const { control } = validateInlineRemovalSource(pkg, part, edit);
  if (edit.before.type === 'inline_image')
    refuseWrite('Tracked image removal needs resource ownership.', 'unsupported-operation');
  const locked = contentControlPropertiesOf(control).lock === 'sdtContentLocked';
  if (locked)
    writer.apply(edit.partName, {
      op: 'setContentControlProperties',
      controlId: control.id,
      lock: 'unlocked',
    });
  writer.apply(edit.partName, {
    op: 'removeContentControl',
    controlId: control.id,
    keepContent: false,
    revision: attribution,
  });
  if (locked)
    writer.edit((working) => {
      const current = working.parts.get(edit.partName)!,
        deleted = findNode(current, control.id);
      if (deleted?.kind !== 'contentControl' || control.kind !== 'contentControl')
        refuseWrite('The deleted control identity changed.', 'apply-failed');
      // Restore the original properties subtree, including native identities and unknown fields.
      const properties = control.children.find(
        (child) => child.namespaceUri === WML_NAMESPACE_URI && child.localName === 'sdtPr'
      );
      if (!properties) refuseWrite('The removed control properties disappeared.', 'apply-failed');
      const restored = replaceNode(current, deleted.id, {
        ...deleted,
        children: deleted.children.map((child) =>
          child.namespaceUri === WML_NAMESPACE_URI && child.localName === 'sdtPr'
            ? properties
            : child
        ),
      });
      if (!restored.ok) refuseWrite('The removed control lock could not restore.', 'apply-failed');
      return withPart(working, restored.part);
    });
}
