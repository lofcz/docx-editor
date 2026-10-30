import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { listNumbering, listPPr, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { setManualList, manualListMarker, manualListState } from './lists';

const context = { documentId: 'list-sequences', epoch: 'mount' };
const numbering = listNumbering.replace(
  '</w:numbering>',
  // Every w:num over one definition shares its counter; a restart makes each a separate list.
  '<w:num w:numId="4"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' +
    '<w:num w:numId="5"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' +
    '<w:num w:numId="6"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>'
);
const heading = (text: string, numId?: string, level = 0, style = 'Heading1', indent = '') =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/>${numId ? listPPr(numId, level) : ''}${indent}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const scenarios = [
  {
    name: 'bridge',
    body:
      listParagraph('One', '1') +
      listParagraph('Gap') +
      listParagraph('Three', '4') +
      listParagraph('Four', '4'),
    at: 1,
    markers: ['1.', '2.', '3.', '4.'],
    moved: [1, 2, 3],
    num: '1',
  },
  {
    name: 'nested item',
    body:
      listParagraph('One', '1') +
      listParagraph('Nested', '2', 1) +
      listParagraph('Gap') +
      listParagraph('Three', '4'),
    at: 2,
    markers: ['1.', '•', '2.', '3.'],
    moved: [2, 3],
    num: '1',
  },
  {
    name: 'different marker',
    body:
      listParagraph('One', '1') +
      listParagraph('Bullet', '2') +
      listParagraph('Gap') +
      listParagraph('Three', '4'),
    at: 2,
    markers: ['1.', '•', '2.', '3.'],
    moved: [2, 3],
    num: '1',
  },
  {
    name: 'nested numbered heading',
    body:
      listParagraph('One', '1') +
      heading('Heading', '5', 1) +
      listParagraph('Gap') +
      listParagraph('Three', '4'),
    at: 2,
    markers: ['1.', '1.', '2.', '3.'],
    moved: [2, 3],
    num: '1',
  },
  {
    name: 'numbered heading boundary',
    body:
      listParagraph('One', '1') +
      heading('Heading', '5') +
      listParagraph('Gap') +
      listParagraph('Two', '4'),
    at: 2,
    markers: ['1.', '1.', '1.', '2.'],
    moved: [2],
    num: '4',
  },
  {
    name: 'plain heading with retained indentation',
    body:
      listParagraph('One', '1') +
      heading('Heading', undefined, 0, 'Heading1', '<w:ind w:left="360"/>') +
      listParagraph('Gap') +
      listParagraph('Three', '4'),
    at: 2,
    markers: ['1.', null, '2.', '3.'],
    moved: [2, 3],
    num: '1',
  },
  {
    name: 'native explicit restart',
    body:
      listParagraph('One', '1') +
      listParagraph('Gap') +
      listParagraph('Five', '6') +
      listParagraph('Six', '6'),
    at: 1,
    markers: ['1.', '2.', '5.', '6.'],
    moved: [1],
    num: '1',
  },
] as const;
const inherited = {
  name: 'inherited numbered heading boundary',
  body:
    listParagraph('One', '1') +
    heading('Heading', '5', 0, 'ImportedTitle') +
    listParagraph('Gap') +
    listParagraph('Two', '4'),
  at: 2,
  markers: ['1.', '1.', '1.', '2.'],
  moved: [2],
  num: '4',
  styles: `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="TitleBase"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ImportedTitle"><w:basedOn w:val="TitleBase"/></w:style></w:styles>`,
};
for (const item of [...scenarios, inherited])
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`${item.name} preserves native sequence boundaries, definitions, full Word trees and history: ${decision}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(item.body, {
          numbering,
          ...('styles' in item ? { styles: item.styles } : {}),
        }),
        author: 'User',
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      try {
        const s = editor.surface!,
          ids = s.session.paragraphIds();
        const point = { paragraphId: ids[item.at]!, offset: 1 };
        s.setSelection({ anchor: point, head: point });
        const initial = s.session.part();
        const before = canonicalOoxmlFingerprint(initial);
        const definition = canonicalOoxmlFingerprint(
          s.session.currentPackage().parts.get('/word/numbering.xml')!
        );
        let expected = initial;
        for (const index of item.moved) {
          const applied = applyTreeOp(expected, {
            op: 'setListNumbering',
            paragraphId: ids[index]!,
            numId: item.num,
            level: 0,
          });
          if (!applied.ok) throw new Error(applied.reason);
          expected = applied.part;
        }
        if (decision !== 'direct')
          expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        const result = setManualList(editor, 'decimal', context);
        if (!result.success) throw new Error(result.error);
        const after = canonicalOoxmlFingerprint(s.session.part());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
        expect(s.state().selection).toEqual({ anchor: point, head: point });
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
        if (decision !== 'direct') {
          const groups = readReviewGroups(s.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
          expect(
            decide(
              editor,
              groups.groups[0]!.id,
              readEditorRevision(editor, context.documentId, context.epoch)!,
              context
            ).success
          ).toBe(true);
        }
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(
          decision === 'reject' ? before : canonicalOoxmlFingerprint(expected)
        );
        if (decision !== 'reject')
          expect(ids.map((id) => manualListMarker(editor, id)?.text ?? null)).toEqual([
            ...item.markers,
          ]);
        expect(
          canonicalOoxmlFingerprint(s.session.currentPackage().parts.get('/word/numbering.xml')!)
        ).toEqual(definition);
        const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!saved.ok) throw new Error(saved.reason);
        expect(
          canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
        ).toEqual(canonicalOoxmlFingerprint(s.session.part()));
      } finally {
        editor.destroy();
      }
    });

for (const headings of [false, true])
  test(`selected paragraphs form independent heading sequences and retain native formatting: headings=${headings}`, () => {
    const body =
      listParagraph('First') +
      (headings ? heading('Middle') : listParagraph('Middle')) +
      listParagraph('Last', headings ? undefined : '4');
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(body, { numbering }),
    });
    try {
      const s = editor.surface!,
        ids = s.session.paragraphIds();
      s.setSelection({
        anchor: { paragraphId: ids[0]!, offset: 0 },
        head: { paragraphId: ids[headings ? 2 : 1]!, offset: 4 },
      });
      const before = canonicalOoxmlFingerprint(s.session.part());
      expect(setManualList(editor, 'decimal', context).success).toBe(true);
      expect(ids.map((id) => manualListMarker(editor, id)?.text)).toEqual(
        headings ? ['1.', '1.', '1.'] : ['1.', '2.', '3.']
      );
      const groups = ids.map((id) => manualListState(editor, context, id)!.numbering!.numId);
      expect(new Set(groups).size).toBe(headings ? 3 : 1);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    } finally {
      editor.destroy();
    }
  });

test('counter bridges preserve unmatched native level extensions', () => {
  const customized = numbering.replace(
    '<w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num>',
    '<w:num w:numId="4"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:futureLevel w:val="keep"/></w:lvl></w:lvlOverride></w:num>'
  );
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      listParagraph('One', '1') + listParagraph('Gap') + listParagraph('Other', '4'),
      { numbering: customized }
    ),
  });
  try {
    const s = editor.surface!,
      ids = s.session.paragraphIds();
    const definition = canonicalOoxmlFingerprint(
      s.session.currentPackage().parts.get('/word/numbering.xml')!
    );
    const point = { paragraphId: ids[1]!, offset: 1 };
    s.setSelection({ anchor: point, head: point });
    expect(setManualList(editor, 'decimal', context).success).toBe(true);
    expect(ids.map((id) => manualListMarker(editor, id)?.text)).toEqual(['1.', '2.', '1.']);
    expect(manualListState(editor, context, ids[2]!)!.numbering!.numId).toBe('4');
    expect(
      canonicalOoxmlFingerprint(s.session.currentPackage().parts.get('/word/numbering.xml')!)
    ).toEqual(definition);
  } finally {
    editor.destroy();
  }
});

test('counter bridges preserve unknown abstract numbering metadata', () => {
  const original = /<w:abstractNum w:abstractNumId="0">.*?<\/w:abstractNum>/.exec(numbering)![0];
  const copy = original
    .replace('w:abstractNumId="0"', 'w:abstractNumId="99"')
    .replace('</w:abstractNum>', '<w:futureTemplate w:val="keep"/></w:abstractNum>');
  const customized = numbering
    .replace(
      '<w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num>',
      '<w:num w:numId="4"><w:abstractNumId w:val="99"/></w:num>'
    )
    .replace('</w:numbering>', copy + '</w:numbering>');
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      listParagraph('One', '1') + listParagraph('Gap') + listParagraph('Other', '4'),
      { numbering: customized }
    ),
  });
  try {
    const s = editor.surface!,
      ids = s.session.paragraphIds();
    const definition = canonicalOoxmlFingerprint(
      s.session.currentPackage().parts.get('/word/numbering.xml')!
    );
    const point = { paragraphId: ids[1]!, offset: 1 };
    s.setSelection({ anchor: point, head: point });
    expect(setManualList(editor, 'decimal', context).success).toBe(true);
    expect(ids.map((id) => manualListMarker(editor, id)?.text)).toEqual(['1.', '2.', '1.']);
    expect(manualListState(editor, context, ids[2]!)!.numbering!.numId).toBe('4');
    expect(
      canonicalOoxmlFingerprint(s.session.currentPackage().parts.get('/word/numbering.xml')!)
    ).toEqual(definition);
  } finally {
    editor.destroy();
  }
});
