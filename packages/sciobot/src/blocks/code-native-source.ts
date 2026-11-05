import {
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { PayloadStorage, StoredPayload } from './payload-opc-codec';
import { codeRecordShape } from './code-record';
import { codeRevisionText, type CodeSourceView } from './code-revision-text';

/** Only plain native code lines are editable. Imported fields and opaque structures stay intact. */
export function nativeCodeSource(
  _part: OoxmlPart,
  control: OoxmlNode,
  view: CodeSourceView = 'plain'
) {
  return codeRevisionText(control, view);
}

/** Source is a read projection; private records retain language and opaque imported metadata. */
export function projectNativeCodeSource(
  pkg: OoxmlPackage,
  entries: Map<string, StoredPayload>,
  storage: PayloadStorage
): void {
  const ids = new Set(
    [...entries.values()]
      .filter(
        (entry) =>
          entry.decoded.status === 'supported' &&
          entry.decoded.value.version === 2 &&
          codeRecordShape(entry.decoded.value)
      )
      .map((entry) => entry.id)
  );
  if (!ids.size) return;
  const owners = new Map<string, { part: OoxmlPart; control: OoxmlNode } | null>();
  for (const part of pkg.parts.values()) {
    if (!isStoryPart(part)) continue;
    for (const { node: control } of contentControlsIn(part.root)) {
      const tag = contentControlPropertiesOf(control).tag;
      if (!tag?.startsWith(storage.tagPrefix)) continue;
      const id = tag.slice(storage.tagPrefix.length);
      if (ids.has(id)) owners.set(id, owners.has(id) ? null : { part, control });
    }
  }
  for (const id of ids) {
    const record = entries.get(id)!;
    if (record.decoded.status !== 'supported') continue;
    const owner = owners.get(id);
    const source = owner && nativeCodeSource(owner.part, owner.control, 'committed');
    if (!source) continue;
    entries.set(id, {
      ...record,
      decoded: {
        status: 'supported',
        value: {
          ...record.decoded.value,
          attributes: { ...record.decoded.value.attributes, value: source.value },
        },
      },
    });
  }
}
