import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { paragraphFragmentsOf, paragraphFragmentsOfBlocks } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  readOoxmlPackage,
  stylesPartOf,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { checklistTag } from '../registry/adapter-names';
import { prepareNativeDocument } from './prepare-document';
import { materializeTaskAppearanceForWord, taskStoryParagraphs } from './task-export';
import { taskExportStyleSource } from './task-export-styles';
import { runCharacterStyle } from './inline-code';

const marker = (tag = checklistTag()) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:strike w:val="0"/><w:spacing w:val="180"/><w:sz w:val="30"/></w:rPr><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>`;
function pkg(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
function prepared(bytes: Uint8Array) {
  const value = prepareNativeDocument(bytes);
  if (!value.success) throw new Error(value.error);
  return value.document;
}
function appearance(bytes: Uint8Array, module: boolean) {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    modules: module ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
  try {
    return editor
      .surface!.layout()
      .pages.flatMap((page) => [
        ...paragraphFragmentsOf(page),
        ...paragraphFragmentsOfBlocks(page.header?.fragments ?? []),
        ...paragraphFragmentsOfBlocks(page.footer?.fragments ?? []),
      ])
      .flatMap((p) => p.lines.flatMap((line) => line.spans))
      .map((span) => ({ text: span.text, style: span.style }));
  } finally {
    editor.destroy();
  }
}
const textRun = (text: string, properties = '') =>
  `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t>${text}</w:t></w:r>`;

for (const inherited of [false, true])
  test(`detached Word styles match native cascade and restore authored runs: inherited strike ${inherited}`, async () => {
    const styles = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr>${inherited ? '<w:strike/>' : ''}<w:color w:val="112233"/></w:rPr></w:style><w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style><w:style w:type="character" w:styleId="Authored"><w:rPr><w:strike/><w:color w:val="0055AA"/></w:rPr></w:style></w:styles>`;
    const source = prepared(
      docxFixture(
        `<w:p>${marker()}${textRun('Plain')}${textRun('Explicit', '<w:strike w:val="0"/><w:color w:val="CC0000"/>')}${textRun('Character', '<w:rStyle w:val="Authored"/>')}${textRun('Emphasis', '<w:rStyle w:val="Emphasis"/>')}</w:p>`,
        { styles }
      )
    );
    const before = source.slice();
    const expected = appearance(source, true);
    expect(expected.find((span) => span.text === 'Plain')!.style).toMatchObject({
      strike: true,
      color: '606060',
    });
    expect(expected.find((span) => span.text === 'Explicit')!.style).toMatchObject({
      strike: false,
      color: 'CC0000',
    });
    expect(expected.find((span) => span.text === 'Character')!.style).toMatchObject({
      strike: !inherited,
      color: '0055AA',
    });
    const exported = await materializeTaskAppearanceForWord(source);
    expect(source).toEqual(before);
    expect(appearance(exported, false)).toEqual(expected);
    expect(await materializeTaskAppearanceForWord(exported)).toBe(exported);
    const originalParagraph = taskStoryParagraphs(pkg(source))[0]!.paragraph;
    const restored = prepared(exported),
      restoredParagraph = taskStoryParagraphs(pkg(restored))[0]!.paragraph;
    expect(canonicalOoxmlFingerprint(restoredParagraph)).toBe(
      canonicalOoxmlFingerprint(originalParagraph)
    );
    expect(appearance(restored, true)).toEqual(expected);
    const second = prepareNativeDocument(restored);
    expect(second).toMatchObject({ success: true, changed: false });
  });

test('table strike toggles and paragraph style toggles use the native cascade', async () => {
  const styles = `<w:styles xmlns:w="${W}"><w:style w:type="table" w:styleId="Table"><w:rPr><w:strike/><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Toggle"><w:rPr><w:strike/></w:rPr></w:style></w:styles>`;
  const source = prepared(
    docxFixture(
      `<w:tbl><w:tblPr><w:tblStyle w:val="Table"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Toggle"/></w:pPr>${marker()}${textRun('TableTask')}</w:p></w:tc></w:tr></w:tbl>`,
      { styles }
    )
  );
  const expected = appearance(source, true);
  expect(expected.find((span) => span.text === 'TableTask')!.style).toMatchObject({
    strike: true,
    bold: true,
  });
  expect(appearance(await materializeTaskAppearanceForWord(source), false)).toEqual(expected);
});

test('imported style collisions and lookalike metadata remain authored', async () => {
  const foreign =
    '<w:style w:type="character" w:customStyle="1" w:styleId="SciobotTaskRun__A_11"><w:name w:val="Sciobot Completed Task Run"/><w:rPr><w:b/></w:rPr></w:style>';
  const source = prepared(
    docxFixture(
      `<w:p>${marker()}${textRun('Plain')}${textRun('Foreign', '<w:rStyle w:val="SciobotTaskRun__A_11"/>')}</w:p>`,
      { styles: `<w:styles xmlns:w="${W}">${foreign}</w:styles>` }
    )
  );
  const original = pkg(source),
    exported = await materializeTaskAppearanceForWord(source),
    output = pkg(exported);
  expect(taskExportStyleSource(original, 'SciobotTaskRun__A_11')).toBeNull();
  expect(canonicalOoxmlFingerprint(stylesPartOf(output)!.root.children[0]!)).toBe(
    canonicalOoxmlFingerprint(stylesPartOf(original)!.root.children[0]!)
  );
  expect(taskExportStyleSource(output, 'SciobotTaskRun__A_11_1')).toEqual({
    base: undefined,
    strikeToggle: true,
    color: true,
  });
  const restored = pkg(prepared(exported));
  const runs = taskStoryParagraphs(restored)[0]!.paragraph.children.filter(
    (node) => node.kind === 'run'
  );
  expect(runCharacterStyle(runs[1]!)).toBe('SciobotTaskRun__A_11');
});

test('pending review content stays exact and ordinary Word checkboxes stay unchanged', async () => {
  const pending = `<w:p><w:pPr><w:pPrChange w:id="20" w:author="Alice"><w:pPr/></w:pPrChange></w:pPr>${marker()}${textRun('Pending')}</w:p>`;
  const pendingOnly = docxFixture(pending);
  expect(await materializeTaskAppearanceForWord(pendingOnly)).toBe(pendingOnly);
  const source = prepared(
    docxFixture(
      `<w:p>${marker()}${textRun('Settled')}</w:p>${pending}<w:p>${marker('ordinary-word-control')}${textRun('Ordinary')}</w:p>`
    )
  );
  const entries = unzipSync(source);
  entries['custom/unrelated.xml'] = strToU8('<foreign  keep="bytes" />');
  const withForeign = zipSync(entries),
    before = taskStoryParagraphs(pkg(withForeign));
  const exported = await materializeTaskAppearanceForWord(withForeign),
    after = taskStoryParagraphs(pkg(exported));
  expect(canonicalOoxmlFingerprint(after[1]!.paragraph)).toBe(
    canonicalOoxmlFingerprint(before[1]!.paragraph)
  );
  expect(canonicalOoxmlFingerprint(after[2]!.paragraph)).toBe(
    canonicalOoxmlFingerprint(before[2]!.paragraph)
  );
  expect(unzipSync(exported)['custom/unrelated.xml']).toEqual(entries['custom/unrelated.xml']);
  expect(strFromU8(unzipSync(exported)['word/document.xml']!)).toContain('w:pPrChange');
});

test('settled header tasks materialize and restore through the same native cascade', async () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const entries = unzipSync(
    docxFixture(
      `<w:p>${textRun('Body')}</w:p><w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr>`,
      {
        relationships: `<Relationship Id="header" Type="${R}/header" Target="header.xml"/>`,
      }
    )
  );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
    )
  );
  entries['word/header.xml'] = strToU8(
    `<w:hdr xmlns:w="${W}"><w:p>${marker()}${textRun('HeaderTask')}</w:p></w:hdr>`
  );
  const source = prepared(zipSync(entries)),
    expected = appearance(source, true);
  expect(expected.find((span) => span.text === 'HeaderTask')!.style).toMatchObject({
    strike: true,
    color: '606060',
  });
  const exported = await materializeTaskAppearanceForWord(source);
  expect(appearance(exported, false)).toEqual(expected);
  const restored = prepared(exported);
  const header = (bytes: Uint8Array) => pkg(bytes).parts.get('/word/header.xml')!.root;
  expect(canonicalOoxmlFingerprint(header(restored))).toBe(
    canonicalOoxmlFingerprint(header(source))
  );
});

for (const base of [undefined, 'Imported'])
  for (const foreign of [false, true])
    test(`legacy preparation restores only exact owned paragraph styles: ${base ?? 'absent'}, foreign ${foreign}`, () => {
      const id = base === undefined ? 'SciobotTaskDone__A' : 'SciobotTaskDone__SImported';
      const styles = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="Sciobot Completed Task"/>${base ? '<w:basedOn w:val="Imported"/>' : ''}<w:semiHidden/><w:rPr><w:strike/><w:color w:val="606060" w:themeColor="text1" w:themeTint="60"/>${foreign ? '<w:i/>' : ''}</w:rPr></w:style></w:styles>`;
      const source = docxFixture(
        `<w:p><w:pPr><w:pStyle w:val="${id}"/></w:pPr>${marker()}${textRun('Legacy')}</w:p>`,
        { styles }
      );
      const result = pkg(prepared(source)),
        paragraph = taskStoryParagraphs(result)[0]!;
      expect(
        directParagraphProperties(paragraph.part, paragraph.paragraph.id).find(
          (property) => property.localName === 'pStyle'
        )?.attributes?.val
      ).toBe(foreign ? id : base);
      expect(canonicalOoxmlFingerprint(stylesPartOf(result)!.root.children[0]!)).toBe(
        canonicalOoxmlFingerprint(stylesPartOf(pkg(source))!.root.children[0]!)
      );
    });
