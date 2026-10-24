import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { paragraphFragmentsOf } from '@docx-editor.dev/core/layout';
import { docxFixture } from '../../test/support/document';
import { listNumbering } from '../../test/support/list-document';
import reference from '../__tests__/fixtures/native-commands.json';
import {
  PARAGRAPH_ALIGNMENTS,
  setParagraphAlignment,
  changeParagraphIndent,
} from './paragraph-position';

const context = { documentId: 'position', epoch: 'mount' };
const paragraph = (rtl: boolean, listed: boolean) =>
  `<w:p><w:pPr>${listed ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' : ''}${rtl ? '<w:bidi/>' : ''}<w:spacing w:line="251" w:lineRule="exact"/><w:ind w:left="720" w:right="123" w:hanging="180"/><w:customKeep w:val="yes"/></w:pPr><w:r><w:t>Alpha beta</w:t></w:r></w:p>`;
for (const rtl of [false, true])
  for (const align of PARAGRAPH_ALIGNMENTS)
    test(`alignment ${align} preserves imported/list properties and selection, one undo, and reopen (${rtl ? 'RTL' : 'LTR'})`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(paragraph(rtl, true) + '<w:p><w:r><w:t>Tail</w:t></w:r></w:p>', {
          numbering: listNumbering,
        }),
      });
      try {
        const s = editor.surface!,
          id = s.session.paragraphIds()[0]!;
        const before = canonicalOoxmlFingerprint(s.session.part());
        const properties = directParagraphProperties(s.session.part(), id);
        const selection = {
          anchor: { paragraphId: id, offset: 5 },
          head: { paragraphId: id, offset: 0 },
        };
        s.setSelection(selection);
        expect(reference.cases.find((c) => c.id === `align.${align}`)).toBeDefined();
        expect(setParagraphAlignment(editor, align, context).success).toBe(true);
        expect(s.state().selection).toEqual(selection);
        expect(
          directParagraphProperties(s.session.part(), id).filter((p) => p.localName !== 'jc')
        ).toEqual([...properties]);
        const expected =
          align === 'justify'
            ? 'both'
            : align === 'start'
              ? rtl
                ? 'right'
                : 'left'
              : align === 'end'
                ? rtl
                  ? 'left'
                  : 'right'
                : align;
        expect(s.formatting().alignment).toBe(expected);
        const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!saved.ok) throw new Error(saved.reason);
        expect(
          canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
        ).toEqual(canonicalOoxmlFingerprint(s.session.part()));
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      } finally {
        editor.destroy();
      }
    });

for (const direction of ['increase', 'decrease'] as const)
  for (const listed of [false, true])
    test(`${direction} changes each paragraph by the production pixel step and keeps list membership (${listed})`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(
          paragraph(false, listed) +
            paragraph(false, false).replace('w:left="720"', 'w:left="360"'),
          { numbering: listNumbering }
        ),
      });
      try {
        const s = editor.surface!,
          ids = s.session.paragraphIds();
        const before = canonicalOoxmlFingerprint(s.session.part());
        const props = ids.map((id) => directParagraphProperties(s.session.part(), id));
        s.setSelection({
          anchor: { paragraphId: ids[0]!, offset: 0 },
          head: { paragraphId: ids[1]!, offset: 5 },
        });
        const fixture = reference.cases.find((c) => c.id === `indent.${direction}`)!;
        expect(fixture).toBeDefined();
        const delta = direction === 'increase' ? 360 : -360;
        expect(changeParagraphIndent(editor, direction, context).success).toBe(true);
        for (const [index, id] of ids.entries()) {
          const after = directParagraphProperties(s.session.part(), id);
          expect(after.filter((p) => p.localName !== 'ind')).toEqual(
            props[index]!.filter((p) => p.localName !== 'ind')
          );
          expect(after.find((p) => p.localName === 'ind')?.attributes).toEqual({
            left: String((index ? 360 : 720) + delta),
            right: '123',
            hanging: '180',
          });
          expect(
            s
              .layout()
              .pages.flatMap((p) => paragraphFragmentsOf(p))
              .find((p) => p.range.paragraphId === id)?.indent.left
          ).toBe(((index ? 360 : 720) + delta) / 20);
        }
        const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!saved.ok) throw new Error(saved.reason);
        expect(
          canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
        ).toEqual(canonicalOoxmlFingerprint(s.session.part()));
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      } finally {
        editor.destroy();
      }
    });
