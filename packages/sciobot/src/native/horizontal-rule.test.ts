import { openDocumentForExport } from '@docx-editor.dev/core/export';
import { openHeadlessDocument } from '@docx-editor.dev/core/store';
import { readWorkingView } from '../review/working-view';
import { readCommittedView } from '../review/committed-view';
import { projectMarkdown } from '../projection/markdown';
import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  stylesPartOf,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import {
  HORIZONTAL_RULE_PARAGRAPH_XML as RULE,
  SCIOBOT_HORIZONTAL_RULE_STYLE_ID as ID,
  horizontalRuleStyleState,
  horizontalRuleProjectionStyles,
  isHorizontalRuleParagraph,
  withHorizontalRuleStyle,
} from './horizontal-rule';
const open = (bytes: Uint8Array) => {
  const result = readOoxmlPackage(bytes);
  if (!result.ok) throw new Error(result.reason);
  return result.package;
};
const paragraph = (pkg: ReturnType<typeof open>) => {
  const body = bodyStoryRoot(pkg.parts.get(pkg.mainDocumentPart)!);
  if (!body || body.kind === 'textValue') throw new Error('Missing body');
  return body.children[0]!;
};

test('the explicit rule identity and border survive native serialization without body writes during preparation', () => {
  const original = open(docxFixture(RULE));
  const body = original.parts.get(original.mainDocumentPart)!;
  const fingerprint = canonicalOoxmlFingerprint(body);
  expect(horizontalRuleStyleState(original)).toBe('missing');
  const prepared = withHorizontalRuleStyle(original)!;
  expect(prepared).not.toBeNull();
  expect(prepared.parts.get(prepared.mainDocumentPart)).toBe(body);
  expect(canonicalOoxmlFingerprint(body)).toBe(fingerprint);
  expect(horizontalRuleStyleState(prepared)).toBe('present');
  expect(withHorizontalRuleStyle(prepared)).toBe(prepared);
  expect(isHorizontalRuleParagraph(paragraph(prepared))).toBe(true);
  const reopened = open(writeOoxmlPackage(prepared));
  expect(horizontalRuleStyleState(reopened)).toBe('present');
  expect(isHorizontalRuleParagraph(paragraph(reopened))).toBe(true);
  expect(canonicalOoxmlFingerprint(reopened.parts.get(reopened.mainDocumentPart)!)).toBe(
    fingerprint
  );
});

test('imported style collisions and duplicates refuse without changing existing metadata', () => {
  for (const definitions of [
    `<w:style w:type="character" w:styleId="${ID}"><w:name w:val="Foreign"/><w:rPr><w:b/></w:rPr></w:style>`,
    `<w:style w:type="paragraph" w:styleId="${ID}"><w:name w:val="Sciobot Horizontal Rule"/><w:next w:val="Normal"/><w:pPr><w:keepNext/></w:pPr></w:style>`,
    `<w:style w:type="paragraph" w:styleId="${ID}"/><w:style w:type="paragraph" w:styleId="${ID}"/>`,
  ]) {
    const pkg = open(
      docxFixture(RULE, { styles: `<w:styles xmlns:w="${W}">${definitions}</w:styles>` })
    );
    const before = canonicalOoxmlFingerprint(stylesPartOf(pkg)!);
    expect(horizontalRuleStyleState(pkg)).toBe('conflict');
    expect(withHorizontalRuleStyle(pkg)).toBeNull();
    expect(canonicalOoxmlFingerprint(stylesPartOf(pkg)!)).toBe(before);
  }
});

test('generic imported borders, nonempty paragraphs, and changed edges are not rule markers', () => {
  for (const xml of [
    RULE.replace(`<w:pStyle w:val="${ID}"/>`, ''),
    RULE.replace('</w:p>', '<w:r><w:t>Keep me</w:t></w:r></w:p>'),
    RULE.replace('w:val="single"', 'w:val="none"'),
    RULE.replace('</w:pBdr>', '<w:top w:val="single" w:sz="6"/></w:pBdr>'),
  ])
    expect(isHorizontalRuleParagraph(paragraph(open(docxFixture(xml))))).toBe(false);
});

test('adding the style preserves imported style nodes and opaque properties by identity', () => {
  const pkg = open(
    docxFixture(RULE, {
      styles: `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Imported"><w:name w:val="Keep"/><w:pPr><w:keepNext/></w:pPr></w:style></w:styles>`,
    })
  );
  const old = stylesPartOf(pkg)!.root.children[0];
  const result = withHorizontalRuleStyle(pkg)!;
  expect(stylesPartOf(result)!.root.children[0]).toBe(old);
  expect(stylesPartOf(result)!.root.children).toHaveLength(2);
});

test('private projection validates native identity and content before emitting generated syntax', async () => {
  for (const malformed of ['', '<w:r><w:fldChar w:fldCharType="begin"/></w:r>', '<w:unknown/>']) {
    const pkg = withHorizontalRuleStyle(
      open(docxFixture(RULE + (malformed ? RULE.replace('</w:p>', malformed + '</w:p>') : '')))
    )!;
    const native = openHeadlessDocument(writeOoxmlPackage(pkg));
    if (!native.ok) throw new Error(native.reason);
    const exported = openDocumentForExport(native.view);
    if (!exported.ok) throw new Error(exported.reason);
    try {
      for (const mode of ['original', 'proposed'] as const) {
        const layout = await exported.session.layoutFor(mode);
        const revision = { documentId: 'rule', epoch: 'rule-mount', revision: 0 };
        const ordinary = projectMarkdown(layout, revision);
        expect(ordinary.text).not.toContain('---');
        const projection = projectMarkdown(layout, revision, { nativePackage: pkg });
        if (malformed) {
          expect(horizontalRuleProjectionStyles(pkg)).toEqual([]);
          expect(projection.text).not.toContain('---');
        } else {
          expect(projection.text).toBe('---');
          expect(projection.sourceMap.sources).toEqual([]);
          expect(projection.sourceMap.paragraphs[0]).toMatchObject({
            sourceStart: 0,
            sourceEnd: 0,
            markdownStart: 0,
            markdownEnd: 3,
          });
          expect(projection.syntax).toEqual([{ start: 0, end: 3, role: 'generated' }]);
        }
      }
    } finally {
      exported.session.dispose();
    }
  }
});

test('unresolved styles relationships and behavior-changing style attributes refuse without writes', () => {
  const unresolved = open(
    docxFixture(RULE, {
      relationships:
        '<Relationship Id="opaqueStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="missing.xml"/>',
    })
  );
  expect(horizontalRuleStyleState(unresolved)).toBe('conflict');
  expect(withHorizontalRuleStyle(unresolved)).toBeNull();
  const prepared = withHorizontalRuleStyle(open(docxFixture(RULE)))!;
  const imported = open(
    docxFixture(RULE, {
      styles: `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:customStyle="1" w:default="1" w:styleId="${ID}"><w:name w:val="Sciobot Horizontal Rule"/><w:next w:val="Normal"/></w:style></w:styles>`,
    })
  );
  expect(horizontalRuleStyleState(prepared)).toBe('present');
  expect(horizontalRuleStyleState(imported)).toBe('conflict');
  expect(horizontalRuleProjectionStyles(imported)).toEqual([]);
  expect(withHorizontalRuleStyle(imported)).toBeNull();
});

test('committed and working readers forward the same immutable package identity', async () => {
  const pkg = withHorizontalRuleStyle(open(docxFixture(RULE)))!;
  const native = openHeadlessDocument(writeOoxmlPackage(pkg));
  if (!native.ok) throw new Error(native.reason);
  const revision = {
    documentId: 'rule-review',
    epoch: 'mount',
    revision: native.view.packageRevision(),
  };
  const before = native.view.currentPackage();
  for (const read of [readCommittedView, readWorkingView]) {
    const result = await read(native.view, revision);
    if (!result.success) throw new Error(result.error);
    expect(result.view.projection.text).toBe('---');
    expect(result.view.revision).toEqual(revision);
    expect(result.view.projection.sourceMap.paragraphs[0]).toMatchObject({
      sourceStart: 0,
      sourceEnd: 0,
      markdownStart: 0,
      markdownEnd: 3,
    });
    expect(native.view.currentPackage()).toBe(before);
  }
});
