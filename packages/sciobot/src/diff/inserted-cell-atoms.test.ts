import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  writeOoxmlPackage,
  withPart,
  withRelationship,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { table } from '../../test/support/target-table-fixture';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { buildBlockIndex } from '../blocks/block-index';
import { stageTrackedPlan } from '../review/tracked-stage';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from './plan-target';
import { createEditPlan, decodeEditPlan, serializeEditPlan } from './plan';
import { insertedAtomsOfEdit } from './inserted-cell-atoms';

const context = { documentId: 'cell-atoms', epoch: 'native' };
const atoms =
  '$x$ ![Cat](asset:cat "Picture") [Ada](mention:user) <source title="Paper" url="https://example.test" />';
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decode = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => /(^|\/)media\//.test(name));
const wordParts = (pkg: OoxmlPackage) =>
  [...pkg.parts]
    .filter(([name]) => name.startsWith('/word/') || name === pkg.mainDocumentPart)
    .map(([name, part]) => [name, canonicalOoxmlFingerprint(part)])
    .sort(([a], [b]) => a!.localeCompare(b!));
const wordRelationships = (pkg: OoxmlPackage) =>
  [...pkg.relationships]
    .filter(([name]) => name.startsWith('/word/') || name === pkg.mainDocumentPart)
    // The parser's order counter is not an OOXML attribute. Keep all persisted relationship fields.
    .map(
      ([name, entries]) => [name, entries.map(({ order: _order, ...fields }) => fields)] as const
    )
    .sort(([a], [b]) => a.localeCompare(b));
const td = (text: string, props = '') => `<td${props}>\n\n${text}\n\n</td>`;
const html = (rows: string[][]) =>
  '<table>\n' + rows.map((row) => '<tr>\n' + row.join('\n') + '\n</tr>').join('\n') + '\n</table>';
const merged =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Peer</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const cases = [
  {
    name: 'new row',
    body: table([
      ['A', 'B'],
      ['Keep', 'Peer'],
    ]),
    target: `| A | B |\n| - | - |\n| Keep | Peer |\n| ${atoms} | Tail |`,
    kind: 'insert-row',
  },
  {
    name: 'new row with a cell list',
    body: table([
      ['A', 'B'],
      ['Keep', 'Peer'],
    ]),
    target: html([
      [td('A'), td('B')],
      [td('Keep'), td('Peer')],
      [td('- ' + atoms + '\n- Plain'), td('Tail')],
    ]),
    kind: 'insert-row',
  },
  {
    name: 'new column',
    body: table([
      ['A', 'B'],
      ['Keep', 'Peer'],
    ]),
    target: `| A | New | B |\n| - | - | - |\n| Keep | ${atoms} | Peer |`,
    kind: 'insert-column',
  },
  {
    name: 'new row and column',
    body: table([
      ['A', 'B'],
      ['Keep', 'Peer'],
    ]),
    target: `| A | New | B |\n| - | - | - |\n| Keep | ${atoms} | Peer |\n| Tail | Plain | End |`,
    kind: 'insert-column',
  },
  {
    name: 'split cell with inserted row',
    body: merged,
    target: html([
      [td('A'), td(atoms)],
      [td('Keep'), td('Peer')],
      [td('Tail'), td('End')],
    ]),
    kind: 'table-geometry',
  },
  {
    name: 'split cell',
    body: merged,
    target: html([
      [td('A'), td(atoms)],
      [td('Keep'), td('Peer')],
    ]),
    kind: 'split-cell',
  },
];
async function source(editor: DocxEditorInstance) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  return {
    revision,
    pkg: editor.surface!.session.currentPackage(),
    projection: view.view.projection,
  };
}
for (const item of cases)
  test(`all four atoms land in an existing table ${item.name}, with native history and Word reopening`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(item.body),
      imageDecodePort: decode,
    });
    try {
      const before = editor.surface!.session.currentPackage();
      const part = editor.surface!.session.part();
      const body = bodyStoryRoot(part)!;
      if (body.kind === 'textValue') throw Error('The document has no body container.');
      const oldTable = body.children.find((node) => node.kind === 'table')!;
      const fingerprint = canonicalOoxmlFingerprint(part);
      const planned = planTargetDocument(await source(editor), item.target, {
        allocateBlockId: (() => {
          let id = 0;
          return () => `cell-atom-${++id}`;
        })(),
      });
      if (!planned.success) throw Error(planned.error);
      const operation = planned.plan.operations.find((edit) => edit.kind === item.kind)!;
      expect(operation).toBeDefined();
      const all = planned.plan.operations.flatMap(insertedAtomsOfEdit);
      expect(all.map((atom) => atom.atom)).toEqual(['math', 'image', 'mention', 'source']);
      expect(
        all.every(
          (atom) => atom.location.kind === 'new-paragraph' && atom.location.path.length >= 2
        )
      ).toBe(true);
      const serialized = serializeEditPlan(planned.plan);
      const roundtrip = decodeEditPlan(serialized);
      if (!roundtrip.success) throw Error(roundtrip.error);
      expect(roundtrip.plan).toEqual(planned.plan);
      let stagedWrites = 0;
      expect(() =>
        stageTrackedPlan(
          {
            currentPackage: () => before,
            apply: () => {
              stagedWrites++;
            },
            edit: () => {
              stagedWrites++;
            },
          },
          roundtrip.plan,
          { author: 'Alice', date: '2026-10-23T12:00:00Z' },
          true
        )
      ).toThrow('Inserted inline atoms need native asset and payload reconciliation.');
      expect(stagedWrites).toBe(0);
      expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(
        writeOoxmlPackage(before)
      );
      const calls: string[] = [];
      const prepared = await prepareInlineImageResources(
        roundtrip.plan,
        async (ref) => {
          calls.push(ref);
          return { bytes: png, mime: 'image/png' };
        },
        decode
      );
      if (!prepared.success) throw Error(prepared.error);
      expect(calls).toEqual(['asset:cat']);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const applied = applyDirectPlan(editor, roundtrip.plan, {
        ...context,
        images: prepared.images,
      });
      stop();
      if (!applied.success) throw Error(applied.error);
      expect(publications).toBe(1);
      expect(findNode(editor.surface!.session.part(), oldTable.id)?.kind).toBe('table');
      const pkg = editor.surface!.session.currentPackage();
      const index = buildBlockIndex(pkg);
      expect(index.diagnostics).toEqual([]);
      expect([...index.byId.values()].map((entry) => entry.payload.type).sort()).toEqual([
        'inline_equation',
        'inline_image',
        'mention',
        'source',
      ]);
      expect(media(pkg)).toHaveLength(1);
      expect(media(pkg)[0]![1]).toEqual(png);
      const again = planTargetDocument(await source(editor), item.target);
      if (!again.success) throw Error(again.error);
      expect(again.plan.operations).toEqual([]);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      // Numbering is a persistent native shell, unlike the image and payload births.
      // Assert its exact retained definitions; every other Word part must restore to baseline.
      let undoBaseline = before;
      if (item.name === 'new row with a cell list') {
        const numbering = pkg.parts.get('/word/numbering.xml')!;
        expect(numbering).toBeDefined();
        expect(
          numbering.root.children.filter(
            (node) => node.kind !== 'textValue' && node.localName === 'abstractNum'
          )
        ).toHaveLength(1);
        expect(
          numbering.root.children.filter(
            (node) => node.kind !== 'textValue' && node.localName === 'num'
          )
        ).toHaveLength(1);
        const related = withRelationship(
          withPart(before, numbering),
          before.mainDocumentPart,
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
          'numbering.xml'
        );
        if (!related.ok) throw Error('The independent numbering shell oracle failed.');
        undoBaseline = related.pkg;
      }

      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
      expect(media(editor.surface!.session.currentPackage())).toEqual([]);
      expect(wordParts(editor.surface!.session.currentPackage())).toEqual(wordParts(undoBaseline));
      expect(wordRelationships(editor.surface!.session.currentPackage())).toEqual(
        wordRelationships(undoBaseline)
      );
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      expect(media(editor.surface!.session.currentPackage())).toEqual(media(pkg));
      const bytes = new Uint8Array(await editor.save());
      const loaded = readOoxmlPackage(bytes);
      if (!loaded.ok) throw Error(loaded.reason);
      expect(media(loaded.package)).toEqual(media(pkg));
      expect(wordParts(loaded.package)).toEqual(wordParts(pkg));
      expect(wordRelationships(loaded.package)).toEqual(wordRelationships(pkg));
      expect(
        [...buildBlockIndex(loaded.package).byId.values()]
          .map((entry) => entry.payload)
          .sort((a, b) => a.id.localeCompare(b.id))
      ).toEqual(
        [...index.byId.values()]
          .map((entry) => entry.payload)
          .sort((a, b) => a.id.localeCompare(b.id))
      );
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
        imageDecodePort: decode,
      });
      try {
        expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(after);
        const replanned = planTargetDocument(await source(reopened), item.target);
        if (!replanned.success) throw Error(replanned.error);
        expect(replanned.plan.operations).toEqual([]);
      } finally {
        reopened.destroy();
      }
      type MutableEdit = {
        kind: string;
        intent: {
          atoms?: Array<{
            offset: number;
            location: { path: number[] };
            resource: { reference: string };
            payload: { id: string };
          }>;
        };
      };
      for (const mutate of [
        (edit: MutableEdit) => {
          delete edit.intent.atoms;
        },
        (edit: MutableEdit) => {
          edit.intent.atoms![0].offset++;
        },
        (edit: MutableEdit) => {
          edit.intent.atoms![0].location.path[0] = 999;
        },
        (edit: MutableEdit) => {
          edit.intent.atoms![1].resource.reference = 'asset:other';
        },
        (edit: MutableEdit) => {
          edit.intent.atoms![2].payload.id = edit.intent.atoms![0].payload.id;
        },
      ]) {
        const altered = JSON.parse(serialized) as { operations: MutableEdit[] };
        mutate(altered.operations.find((edit) => edit.kind === item.kind)!);
        expect(decodeEditPlan(JSON.stringify(altered)).success).toBe(false);
      }
      const duplicate = createEditPlan(roundtrip.plan.revision, [
        ...roundtrip.plan.operations,
        operation,
      ]);
      expect(duplicate.success).toBe(false);
    } finally {
      editor.destroy();
    }
  });
