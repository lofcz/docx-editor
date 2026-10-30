import { expect, test } from 'bun:test';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  readOoxmlPart,
  writeOoxmlPackage,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { propertyEditor } from '../../test/support/property-execution';
import { listDocument, listParagraph } from '../../test/support/list-document';
import { prepareReviewGroupDocument } from '../native/prepare-review-groups';
import { planListMembership } from '../diff/lists';
import { readEditorRevision } from '../projection/revision-cache';
import { trackFormatPlan } from '../review/track-format';
import { applyDirectPlan } from '../ai/apply-direct';
import { readReviewGroupStore, withReviewGroupRecord } from '../review/group-record-store';
import { encodeReviewGroupRecord, type ReviewGroupRecord } from '../review/group-record-schema';
const context = {
  documentId: 'format-package',
  epoch: 'mount',
  author: 'Alice',
  date: '2026-10-23T00:00:00Z',
};
function reopen(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return read.package;
}
function shape(input: OoxmlPackage) {
  const pkg = reopen(input);
  return {
    parts: [...pkg.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]).sort(),
    bytes: [...pkg.partBytes]
      .filter(([name]) => !pkg.parts.has(name))
      .map(([name, bytes]) => [name, [...bytes]])
      .sort(),
    types: {
      defaults: [...pkg.contentTypes.defaults].sort(),
      overrides: [...pkg.contentTypes.overrides].sort(),
    },
    rels: [...pkg.relationships].map(([name, rels]) => [name, rels]).sort(),
    external: [...pkg.externalTargets].sort(),
    main: pkg.mainDocumentPart,
  };
}
for (const fresh of [false, true])
  for (const action of ['accept', 'reject'] as const)
    test(`complete ${fresh ? 'allocated' : 'existing'} numbering ${action} package matches an independent expected ledger`, async () => {
      const entries = unzipSync(listDocument(listParagraph('Item')));
      entries['word/media/preserved.bin'] = new Uint8Array([0, 1, 2, 127, 128, 254, 255]);
      entries['opaque/preserved.xml'] = strToU8(
        '<x:opaque xmlns:x="urn:r06:opaque" x:stamp="keep"><x:child>original</x:child></x:opaque>'
      );
      entries['word/header1.xml'] = strToU8(
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Keep header</w:t></w:r></w:p></w:hdr>'
      );
      entries['[Content_Types].xml'] = strToU8(
        strFromU8(entries['[Content_Types].xml']!).replace(
          '</Types>',
          '<Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
        )
      );
      const prep = prepareReviewGroupDocument(zipSync(entries));
      if (!prep.success) throw Error(prep.error);
      const read = readOoxmlPackage(prep.document);
      if (!read.ok) throw Error(read.reason);
      const editor = propertyEditor(read.package),
        direct = propertyEditor(read.package);
      const initial = direct.surface!.session.currentPackage();
      const plan = (value: typeof editor) => {
        const s = value.surface!.session;
        const p = planListMembership(
          {
            pkg: s.currentPackage(),
            partName: s.part().name,
            revision: readEditorRevision(value, context.documentId, context.epoch)!,
          },
          [
            {
              paragraphId: s.paragraphIds()[0]!,
              target: {
                kind: 'ordered',
                level: 2,
                instance: fresh
                  ? { kind: 'new', key: 'review', start: 5 }
                  : { kind: 'existing', numId: '1' },
              },
            },
          ]
        );
        if (!p.success) throw Error(p.error);
        return p.plan;
      };
      try {
        const a = applyDirectPlan(direct, plan(direct), context);
        if (!a.success) throw Error(a.error);
        const t = trackFormatPlan(editor, plan(editor), context);
        if (!t.success) throw Error(t.error);
        const store = readReviewGroupStore(editor.surface!.session.currentPackage());
        if (store.status !== 'ready') throw Error('Missing group store');
        expect(store.retainedRecords.size).toBe(1);
        const [id, actualRecord] = [...store.retainedRecords][0]!;
        expect(id).toMatch(/^rg_[0-9a-f]{32}$/);
        const expectedRecord: ReviewGroupRecord = {
          schemaVersion: 1,
          id,
          members: [
            {
              partName: '/word/document.xml',
              revision: { id: '1', author: 'Alice', date: '2026-10-23T00:00:00.000Z' },
            },
          ],
        };
        expect(actualRecord.raw).toBe(encodeReviewGroupRecord(expectedRecord));
        const originalNumbering = initial.parts.get('/word/numbering.xml')!;
        if (originalNumbering.root.kind !== 'generic')
          throw Error('Expected generic numbering root');
        let expectedNumbering = originalNumbering;
        if (fresh) {
          const seed = readOoxmlPart(
            // A fresh w:num over a reused definition restarts every level at its own start.
            '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="4"><w:abstractNumId w:val="0"/>' +
              Array.from(
                { length: 9 },
                (_, level) =>
                  `<w:lvlOverride w:ilvl="${level}"><w:startOverride w:val="${level === 2 ? 5 : 1}"/></w:lvlOverride>`
              ).join('') +
              '</w:num></w:numbering>',
            { name: '/r06-resource-oracle.xml', contentType: originalNumbering.contentType }
          );
          if (!seed.ok) throw Error(seed.reason);
          expectedNumbering = {
            ...originalNumbering,
            root: {
              ...originalNumbering.root,
              children: [...originalNumbering.root.children, ...seed.part.root.children],
            },
          };
        }
        expect(
          canonicalOoxmlFingerprint(
            direct.surface!.session.currentPackage().parts.get('/word/numbering.xml')!
          )
        ).toBe(canonicalOoxmlFingerprint(expectedNumbering));
        const chosen =
          action === 'accept'
            ? direct.surface!.session.currentPackage()
            : withPart(initial, expectedNumbering);
        const expected = withReviewGroupRecord(chosen, expectedRecord);
        if (!expected.ok) throw Error(expected.reason);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!saved.ok) throw Error(saved.reason);
        expect(shape(saved.package)).toEqual(shape(expected.pkg));
        expect(saved.package.parts.get('/word/numbering.xml')).toBeDefined();
      } finally {
        editor.destroy();
        direct.destroy();
      }
    });
