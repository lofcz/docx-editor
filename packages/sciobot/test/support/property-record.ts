import { escapeXml, readOoxmlPackage, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import type { BlockPayload } from '../../src/blocks/payload-schema';
import { GENERATED_LABEL_TAG } from '../../src/projection/block-tags';
import { withPropertyReviewRecord } from '../../src/review/property-record-store';
import type { PropertyReviewRecord } from '../../src/review/property-record-schema';
import { docxFixture } from './document';
import { sourceControl, sourceParagraph } from './target-source-fixture';

export const propertyRecordRevision = {
  documentId: 'property-records',
  epoch: 'load-1',
  revision: 0,
};
export function propertyRecordFixture() {
  const label = `<w:sdt><w:sdtPr><w:tag w:val="${escapeXml(GENERATED_LABEL_TAG)}"/></w:sdtPr><w:sdtContent>${sourceParagraph('Old activity — 10 minutes')}</w:sdtContent></w:sdt>`,
    loaded = readOoxmlPackage(
      docxFixture(
        sourceControl('phase', sourceControl('activity', label + sourceParagraph('Body'))) +
          sourceParagraph('Outside') +
          '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
      )
    );
  if (!loaded.ok) throw new Error(loaded.reason);
  let pkg = loaded.package;
  const payloads: BlockPayload[] = [
    {
      id: 'phase',
      type: 'lesson_phase',
      version: 1 as const,
      attributes: { phaseType: 'evocation' },
    },
    {
      id: 'activity',
      type: 'lesson_activity',
      version: 1 as const,
      attributes: {
        name: 'Old activity',
        duration: 10,
        resource: 'keep',
        opaque: { audio: 'keep' },
      },
      extension: { token: 'keep' },
    },
  ];
  for (const payload of payloads) {
    const written = withPayloadRecord(pkg, payload, payload.id + ' label');
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return pkg;
}
export function propertyRecord(id = 'proposal-one'): PropertyReviewRecord {
  return {
    schemaVersion: 1,
    id,
    author: 'ScioBot',
    date: '2026-10-19T09:00:00.000Z',
    predecessors: [],
    blocks: [
      {
        blockId: 'activity',
        type: 'lesson_activity',
        changes: [
          {
            name: 'duration',
            before: { present: true, value: 10 },
            after: { present: true, value: 15 },
          },
        ],
      },
    ],
    nativeRevisions: [
      {
        partName: '/word/document.xml',
        id: '5',
        author: 'ScioBot',
        date: '2026-10-19T09:00:00.000Z',
      },
    ],
    extension: { unrelated: { keep: true } },
  };
}
export function packageWithPropertyRecord(pkg: OoxmlPackage, record = propertyRecord()) {
  const result = withPropertyReviewRecord(pkg, record);
  if (!result.ok) throw new Error(result.reason);
  return result.pkg;
}
export const propertyActivity = (pkg: OoxmlPackage) => buildBlockIndex(pkg).byId.get('activity')!;
