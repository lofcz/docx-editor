import {
  contentControlPropertiesOf,
  lockForbidsEdit,
  parentNodeOf,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';

/** The caller owns its target. Native ancestors retain their independent edit restrictions. */
export function ancestorContentWriteRefusal(part: OoxmlPart, nodeId: string): string | null {
  for (let parent = parentNodeOf(part, nodeId); parent; parent = parentNodeOf(part, parent.id)) {
    if (parent.kind !== 'contentControl') continue;
    const properties = contentControlPropertiesOf(parent);
    if (properties.dataBinding) return 'An enclosing native control is bound to custom XML.';
    if (lockForbidsEdit(properties.lock)) return 'An enclosing native control locks its content.';
  }
  return null;
}
