import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  paragraphTextOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { PayloadStorage, StoredPayload } from './payload-opc-codec';
import { codeRecordShape } from './code-record';
import { DEFAULT_PAYLOAD_LIMITS } from './payload-schema';

/** Only plain native code lines are editable. Imported fields and opaque structures stay intact. */
export function nativeCodeSource(part: OoxmlPart, control: OoxmlNode) {
  const content = contentControlContentNodeOf(control);
  if (!content || !content.children.length || content.children.length > 10000) return null;
  const paragraphs = content.children;
  for (const paragraph of paragraphs) {
    if (paragraph.kind !== 'paragraph') return null;
    for (const child of paragraph.children) {
      if (child.kind === 'paragraphProperties') {
        if (
          child.children.some((property) =>
            ['sectPr', 'pPrChange', 'rPr'].includes(property.localName)
          )
        )
          return null;
        continue;
      }
      if (child.kind !== 'run') return null;
      for (const runChild of child.children) {
        if (runChild.kind === 'runProperties') {
          if (
            runChild.children.some((property) =>
              ['rPrChange', 'vanish', 'webHidden', 'specVanish'].includes(property.localName)
            )
          )
            return null;
          continue;
        }
        if (
          runChild.kind === 'text' &&
          runChild.children.every((value) => value.kind === 'textValue')
        )
          continue;
        if (
          runChild.namespaceUri === WML_NAMESPACE_URI &&
          runChild.localName === 'tab' &&
          !runChild.children.length
        )
          continue;
        return null;
      }
    }
  }
  const lines = paragraphs.map((paragraph) => paragraphTextOf(part, paragraph.id) ?? '');
  const value = lines.join('\n');
  if (
    value.length > DEFAULT_PAYLOAD_LIMITS.maxBytes ||
    new TextEncoder().encode(value).byteLength > DEFAULT_PAYLOAD_LIMITS.maxBytes
  )
    return null;
  return { content, paragraphs, lines, value };
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
    const source = owner && nativeCodeSource(owner.part, owner.control);
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
