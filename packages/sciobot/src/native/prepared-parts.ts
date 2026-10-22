import { serializeOoxmlPart, writeZip, type OoxmlPackage } from '@docx-editor.dev/core/store';

/** Serialize changed infrastructure only. Never run Word export cleanup on imported stories. */
export function initializedPackageBytes(
  document: Uint8Array,
  original: OoxmlPackage,
  prepared: OoxmlPackage
): Uint8Array {
  if (prepared === original) return document;
  const entries = new Map(prepared.partBytes);
  for (const [name, part] of prepared.parts)
    if (part !== original.parts.get(name))
      entries.set(name, new TextEncoder().encode(serializeOoxmlPart(part)));
  return writeZip(entries);
}
