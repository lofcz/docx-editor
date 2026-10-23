import type { BlockPayload } from '../../src/blocks/payload-schema';
import {
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import { createBlockPresenceRecord } from '../../src/review/block-presence-plan';
import { sourceControl, sourceParagraph } from './target-source-fixture';
import { docxFixture } from './document';
export const presenceAttribution = {
  id: 'presence-one',
  author: 'Bot',
  date: '2026-10-23T00:00:00.000Z',
};
export function presencePackage(withBlocks = true) {
  const read = readOoxmlPackage(
    docxFixture(
      (withBlocks
        ? sourceControl('outer', sourceControl('inner', sourceParagraph('Opaque child body')))
        : '') + sourceParagraph('Outside')
    )
  );
  if (!read.ok) throw Error(read.reason);
  let pkg = read.package;
  const payloads: readonly BlockPayload[] = [
    {
      id: 'outer',
      type: 'lesson_phase',
      version: 1 as const,
      attributes: { phaseType: 'evocation', opaque: null },
      extension: { keep: ['root', 7] },
    },
    {
      id: 'inner',
      type: 'custom_atomic',
      version: 1 as const,
      attributes: { unknown: { asset: 'asset:existing', data: ['keep'] } },
      extension: { keep: true },
    },
  ];
  if (withBlocks)
    for (const payload of payloads) {
      const written = withPayloadRecord(pkg, payload, payload.id + ' label');
      if (!written.ok) throw Error(written.reason);
      pkg = written.pkg;
    }
  return pkg;
}
export function presenceRecord(
  before = presencePackage(false),
  canonical = presencePackage(),
  remove = false
) {
  const planned = createBlockPresenceRecord(
    remove ? canonical : before,
    canonical,
    [{ blockId: 'outer', before: { present: remove }, after: { present: !remove } }],
    presenceAttribution
  );
  if (!planned.success) throw Error(planned.error);
  return planned.record;
}
export function reopenPresence(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return read.package;
}
