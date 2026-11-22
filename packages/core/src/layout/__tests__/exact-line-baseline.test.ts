import { expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { applyLineSpacing } from '../paragraph-style.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import {
  createFixedMeasurer,
  createLayoutSession,
  layoutSemanticDocument,
} from '../semantic-layout.ts';
import { createParagraphLayoutCache } from '../layout-cache.ts';
import { compatibilityProfileFromSettings } from '../compatibility/compatibility-profile.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function part(xml: string, name = '/word/settings.xml') {
  const parsed = readOoxmlPart(xml, { name, contentType: 'app/xml' });
  if (!parsed.ok) throw new Error(parsed.reason);
  return parsed.part;
}
function settings(mode: number | undefined, value = '') {
  return part(`<w:settings xmlns:w="${W}" xmlns:x="urn:foreign"><w:compat>
    ${mode === undefined ? '' : `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`}
    ${value}</w:compat></w:settings>`).root;
}

test('exact baselines are face-independent across short and tall authored heights', () => {
  for (const height of [8, 10, 12, 14, 18, 24, 36]) {
    for (const [natural, baseline] of [
      [12.649, 10.318],
      [14, 11],
      [24, 18],
    ]) {
      expect(applyLineSpacing({ rule: 'exact', value: height }, natural!, baseline!)).toEqual({
        height,
        baseline: height * 0.8,
      });
    }
  }
});

test('legacy noExtraLineSpacing retains natural baselines only within the authored height', () => {
  expect(
    applyLineSpacing({ rule: 'exact', value: 24, preserveExactBaseline: true }, 14, 11)
  ).toEqual({ height: 24, baseline: 11 });
  expect(
    applyLineSpacing({ rule: 'exact', value: 8, preserveExactBaseline: true }, 14, 11)
  ).toEqual({ height: 8, baseline: 8 });
  // Every mode keeps the face baseline, also mode 15 and a refused mode.
  for (const mode of [undefined, 11, 12, 14, 15, 99]) {
    expect(
      compatibilityProfileFromSettings(settings(mode, '<w:noExtraLineSpacing/>')).has(
        'preserveExactLineBaseline'
      )
    ).toBe(true);
    for (const value of ['0', 'false', 'off', 'invalid'])
      expect(
        compatibilityProfileFromSettings(
          settings(mode, `<w:noExtraLineSpacing w:val="${value}"/>`)
        ).has('preserveExactLineBaseline')
      ).toBe(false);
  }
  expect(
    compatibilityProfileFromSettings(settings(12, '<x:noExtraLineSpacing/>')).has(
      'preserveExactLineBaseline'
    )
  ).toBe(false);
});

test('baseline policy reaches body and cell lines and invalidates reused paragraph geometry', () => {
  const paragraph =
    '<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>Text</w:t></w:r></w:p>';
  const document = part(
    `<w:document xmlns:w="${W}"><w:body>${paragraph}<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${paragraph}</w:tc></w:tr></w:tbl></w:body></w:document>`,
    '/word/document.xml'
  );
  const session = createLayoutSession();
  const cache = createParagraphLayoutCache();
  const measurer = createFixedMeasurer(6, 14);
  const modern = buildStyleCascadeTable(null, undefined, settings(15));
  const legacy = buildStyleCascadeTable(null, undefined, settings(12, '<w:noExtraLineSpacing/>'));
  expect(modern.cacheToken).not.toBe(legacy.cacheToken);
  for (const cascade of [modern, legacy, modern]) {
    const options = { measurer, styleCascade: cascade };
    const warm = layoutSemanticDocument(document, 1, { ...options, session, cache });
    const cold = layoutSemanticDocument(document, 1, options);
    expect(warm.pages).toEqual(cold.pages);
    const body = warm.pages[0]!.fragments.find((block) => block.kind === 'paragraph')!;
    const table = warm.pages[0]!.fragments.find((block) => block.kind === 'table')!;
    const cell = table.rows[0]!.cells[0]!.blocks.find((block) => block.kind === 'paragraph')!;
    for (const line of [body.lines[0]!, cell.lines[0]!]) {
      expect(line.box.height).toBe(20);
      expect(line.baseline).toBeCloseTo(cascade === legacy ? 11.2 : 16, 8);
    }
  }
});

test('the engine cssHalfLeading setting centres the text band in exact and at-least lines', () => {
  const engine = (val = '1') =>
    settings(
      15,
      `<w:compatSetting w:name="cssHalfLeading" w:uri="urn:docx-editor.dev:layout" w:val="${val}"/>`
    );
  expect(compatibilityProfileFromSettings(engine()).has('cssHalfLeading')).toBe(true);
  expect(compatibilityProfileFromSettings(engine('0')).has('cssHalfLeading')).toBe(false);
  expect(compatibilityProfileFromSettings(settings(15)).has('cssHalfLeading')).toBe(false);
  // A Word URI does not turn it on.
  expect(
    compatibilityProfileFromSettings(
      settings(
        15,
        '<w:compatSetting w:name="cssHalfLeading" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
      )
    ).has('cssHalfLeading')
  ).toBe(false);
  // 19.5 pt line, 14.5 pt text band with an 11.6 pt ascent: 2.5 pt above and below.
  expect(applyLineSpacing({ rule: 'atLeast', value: 19.5, halfLeading: true }, 14.5, 11.6)).toEqual(
    {
      height: 19.5,
      baseline: 14.1,
      trailing: 2.5,
    }
  );
  expect(applyLineSpacing({ rule: 'exact', value: 19.5, halfLeading: true }, 14.5, 11.6)).toEqual({
    height: 19.5,
    baseline: 14.1,
    trailing: 2.5,
  });
  // A shorter exact line overlaps the band evenly; an at-least value below the text is natural.
  expect(applyLineSpacing({ rule: 'exact', value: 12.5, halfLeading: true }, 14.5, 11.6)).toEqual({
    height: 12.5,
    baseline: 10.6,
    trailing: 0,
  });
  expect(applyLineSpacing({ rule: 'atLeast', value: 10, halfLeading: true }, 14.5, 11.6)).toEqual({
    height: 14.5,
    baseline: 11.6,
    trailing: 0,
  });
  // Without the setting Word's placement stays: at-least extra above, exact 80/20.
  expect(applyLineSpacing({ rule: 'atLeast', value: 19.5 }, 14.5, 11.6).baseline).toBeCloseTo(16.6);
});
