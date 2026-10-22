import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { compileInlineTarget } from '../diff/inline-target';
import { parseTargetMarkdown } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from '../review/committed-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import {
  applyInlineCode,
  inlineCodeOps,
  inlineCodeStyleState,
  inlineCodeStyleId,
  isInlineCodeRun,
  SCIOBOT_CODE_STYLE_ID,
  withInlineCodeStyle,
  type InlineCodeRequest,
} from './inline-code';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const context = { documentId: 'inline-code', epoch: 'mount-1' };
const STYLES =
  `<w:styles xmlns:w="${W}">` +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
  '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>' +
  '</w:styles>';

const containers: HTMLElement[] = [];
function mount(document_: Uint8Array): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  return createDocxEditor({
    container,
    document: document_,
    author: 'Teacher',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
}
const token = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, context.documentId, context.epoch)!;
const fingerprint = (editor: DocxEditorInstance) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part());
const paragraphIds = (editor: DocxEditorInstance) => editor.surface!.session.paragraphIds();
async function working(editor: DocxEditorInstance): Promise<string> {
  const view = await readWorkingView(editor.surface!.session, token(editor));
  if (!view.success) throw new Error(view.error);
  return view.view.projection.text;
}
async function committed(editor: DocxEditorInstance): Promise<string> {
  const view = await readCommittedView(editor.surface!.session, token(editor));
  if (!view.success) throw new Error(view.error);
  return view.view.projection.text;
}
function apply(editor: DocxEditorInstance, request: InlineCodeRequest) {
  const result = applyInlineCode(editor, request, { ...context, revision: token(editor) });
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result;
}
/** Exact Markdown AST oracle: paragraph text tokens with their semantic marks. */
function inlineRuns(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.ast.children.map((node) => {
    const compiled = compileInlineTarget(node);
    if (!compiled.success) throw new Error(compiled.error);
    return compiled.target.tokens.map((entry) => [
      entry.kind === 'text' ? entry.value : entry.kind,
      [...entry.marks],
    ]);
  });
}
function savedParts(bytes: ArrayBuffer) {
  const entries = unzipSync(new Uint8Array(bytes));
  return Object.fromEntries(
    Object.entries(entries)
      .filter(([name]) => name.endsWith('.xml') || name.endsWith('.rels'))
      .map(([name, value]) => [name, strFromU8(value)])
  );
}

test('code lands in one transaction, creates its style, and undoes in one step', async () => {
  const editor = mount(
    docxFixture(
      '<w:p><w:r><w:rPr><w:b/><w:customKeep w:val="keep"/></w:rPr><w:t xml:space="preserve">alpha beta gamma</w:t></w:r></w:p>'
    )
  );
  try {
    const id = paragraphIds(editor)[0]!,
      before = fingerprint(editor);
    expect(inlineCodeStyleState(editor.surface!.session.currentPackage())).toBe('missing');
    const result = apply(editor, { paragraphId: id, start: 6, end: 10, code: true });
    expect(result.committed).toBe(true);
    expect(await working(editor)).toBe('**alpha `beta` gamma**');
    expect(inlineRuns(await working(editor))).toEqual([
      [
        ['alpha ', ['bold']],
        ['beta', ['bold', 'code']],
        [' gamma', ['bold']],
      ],
    ]);
    const part = editor.surface!.session.part();
    const coded = editor.surface!.session.part().root;
    expect(coded).toBeTruthy();
    const xml = savedParts(await editor.save());
    expect(xml['word/document.xml']).toContain(
      `<w:rPr><w:rStyle w:val="${SCIOBOT_CODE_STYLE_ID}"/><w:b/><w:customKeep w:val="keep"/></w:rPr><w:t>beta</w:t>`
    );
    expect(xml['word/styles.xml']!.match(/w:styleId="SciobotCode"/g)).toHaveLength(1);
    expect(xml['word/_rels/document.xml.rels']).toContain('Target="styles.xml"');
    expect(xml['[Content_Types].xml']).toContain('/word/styles.xml');
    expect(paragraphTextOf(part, id)).toBe('alpha beta gamma');
    expect(editor.getReviewItems()).toEqual([]);

    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
    expect(inlineCodeStyleState(editor.surface!.session.currentPackage())).toBe('missing');
    expect(editor.surface!.session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(await working(editor)).toBe('**alpha `beta` gamma**');

    // Reopen: the explicit style identity survives the DOCX round trip.
    const reopened = mount(new Uint8Array(await editor.save()));
    try {
      expect(await working(reopened)).toBe('**alpha `beta` gamma**');
      expect(inlineCodeStyleState(reopened.surface!.session.currentPackage())).toBe('present');
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('code survives native formatting and typing, and removal restores ordinary text', async () => {
  const editor = mount(
    docxFixture('<w:p><w:r><w:t xml:space="preserve">see code here</w:t></w:r></w:p>')
  );
  try {
    const id = paragraphIds(editor)[0]!,
      surface = editor.surface!;
    apply(editor, { paragraphId: id, start: 4, end: 8, code: true });
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 13 },
    });
    surface.toggleRunProperty('i');
    expect(await working(editor)).toBe('*see `code` here*');
    // Typing inside the code run extends the code span; typing outside stays ordinary.
    expect(
      surface.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: id, offset: 6, text: 'X' },
        { op: 'insertText', paragraphId: id, offset: 0, text: 'Now ' },
      ]).rejected
    ).toBe(false);
    expect(await working(editor)).toBe('*Now see `coXde` here*');
    expect(inlineRuns(await working(editor))[0]).toEqual([
      ['Now see ', ['italic']],
      ['coXde', ['italic', 'code']],
      [' here', ['italic']],
    ]);
    // Removing the middle leaves two code spans with an ordinary character between them.
    apply(editor, { paragraphId: id, start: 10, end: 11, code: false });
    expect(await working(editor)).toBe('*Now see `co`X`de` here*');
    // Adjacent code runs with one style set project as one code span again.
    apply(editor, { paragraphId: id, start: 8, end: 13, code: true });
    expect(await working(editor)).toBe('*Now see `coXde` here*');
    apply(editor, { paragraphId: id, start: 8, end: 13, code: false });
    expect(await working(editor)).toBe('*Now see coXde here*');
    for (const node of surface.session.part().root.children)
      expect(isInlineCodeRun(node)).toBe(false);
    expect(savedParts(await editor.save())['word/document.xml']).not.toContain('SciobotCode');
  } finally {
    editor.destroy();
  }
});

test('backticks, XML characters, and surrogate pairs keep exact text and addresses', async () => {
  const value = 'run `a<b>&c` 😀 ``';
  const editor = mount(
    docxFixture(
      `<w:p><w:r><w:t xml:space="preserve">x ${value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')} y</w:t></w:r></w:p>`
    )
  );
  try {
    const id = paragraphIds(editor)[0]!;
    expect(
      inlineCodeOps(editor.surface!.session.currentPackage(), {
        paragraphId: id,
        start: 2,
        end: 2 + value.indexOf('😀') + 1,
        code: true,
      })
    ).toMatchObject({ success: false, code: 'ambiguous-target' });
    apply(editor, { paragraphId: id, start: 2, end: 2 + value.length, code: true });
    const markdown = await working(editor);
    expect(markdown).toBe('x ``` run `a<b>&c` 😀 `` ``` y');
    expect(inlineRuns(markdown)[0]).toEqual([
      ['x ', []],
      [value, ['code']],
      [' y', []],
    ]);
    expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(`x ${value} y`);
  } finally {
    editor.destroy();
  }
});

test('links keep their target, and link text returns to the Hyperlink style', async () => {
  const editor = mount(
    docxFixture(
      '<w:p><w:bookmarkStart w:id="1" w:name="here"/><w:r><w:t xml:space="preserve">Go </w:t></w:r>' +
        '<w:hyperlink w:anchor="here"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>there</w:t></w:r></w:hyperlink>' +
        '<w:bookmarkEnd w:id="1"/></w:p>',
      { styles: STYLES }
    )
  );
  try {
    const id = paragraphIds(editor)[0]!;
    apply(editor, { paragraphId: id, start: 3, end: 8, code: true });
    expect(await working(editor)).toBe('Go [`there`](#here)');
    apply(editor, { paragraphId: id, start: 3, end: 8, code: false });
    expect(await working(editor)).toBe('Go [there](#here)');
    expect(savedParts(await editor.save())['word/document.xml']).toContain(
      '<w:hyperlink w:anchor="here"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>there</w:t></w:r></w:hyperlink>'
    );
  } finally {
    editor.destroy();
  }
});

test('code retains imported character styles and unknown run properties through removal and reopen', async () => {
  const editor = mount(
    docxFixture(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t xml:space="preserve">mono </w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="Strong"/><w:customKeep w:val="keep"/></w:rPr><w:t>strong</w:t></w:r></w:p>',
      { styles: STYLES }
    )
  );
  try {
    const id = paragraphIds(editor)[0]!;
    expect(await working(editor)).toBe('mono **strong**');
    const before = fingerprint(editor);
    const originalStyles = savedParts(await editor.save())['word/styles.xml']!;
    apply(editor, { paragraphId: id, start: 0, end: 11, code: true });
    expect(inlineRuns(await working(editor))[0]).toEqual([
      ['mono ', ['code']],
      ['strong', ['bold', 'code']],
    ]);
    const saved = await editor.save();
    const parts = savedParts(saved);
    expect(parts['word/document.xml']).toContain(
      `<w:rStyle w:val="${inlineCodeStyleId('Strong')}"/>`
    );
    expect(parts['word/document.xml']).toContain('<w:customKeep w:val="keep"/>');
    expect(parts['word/document.xml']).toContain('w:ascii="Courier New"');
    expect(parts['word/styles.xml']).toContain('<w:basedOn w:val="Strong"/>');
    // The imported style definition is unchanged; code inherits rather than copying its marks.
    const imported = /<w:style [^>]*w:styleId="Strong"[\s\S]*?<\/w:style>/;
    expect(parts['word/styles.xml']!.match(imported)![0]).toBe(originalStyles.match(imported)![0]);
    const reopened = mount(new Uint8Array(saved));
    try {
      expect(inlineRuns(await working(reopened))).toEqual(inlineRuns(await working(editor)));
      apply(reopened, { paragraphId: paragraphIds(reopened)[0]!, start: 0, end: 11, code: false });
      expect(await working(reopened)).toBe('mono **strong**');
      expect(savedParts(await reopened.save())['word/document.xml']).toContain(
        '<w:rStyle w:val="Strong"/>'
      );
    } finally {
      reopened.destroy();
    }
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    apply(editor, { paragraphId: id, start: 0, end: 11, code: false });
    expect(await working(editor)).toBe('mono **strong**');
    expect(savedParts(await editor.save())['word/document.xml']).toContain(
      '<w:customKeep w:val="keep"/>'
    );
  } finally {
    editor.destroy();
  }
});

test('encoded imported style IDs restore exactly and conflicting derived styles refuse atomically', async () => {
  const base = 'Style & český / 😀';
  const escaped = 'Style &amp; český / 😀';
  const style = `<w:style w:type="character" w:styleId="${escaped}"><w:name w:val="Imported"/></w:style>`;
  const body = `<w:p><w:r><w:rPr><w:rStyle w:val="${escaped}"/></w:rPr><w:t>text</w:t></w:r></w:p>`;
  const editor = mount(
    docxFixture(body, { styles: `<w:styles xmlns:w="${W}">${style}</w:styles>` })
  );
  try {
    const paragraphId = paragraphIds(editor)[0]!;
    apply(editor, { paragraphId, start: 0, end: 4, code: true });
    expect(await working(editor)).toBe('`text`');
    apply(editor, { paragraphId, start: 0, end: 4, code: false });
    expect(await working(editor)).toBe('text');
    expect(savedParts(await editor.save())['word/document.xml']).toContain(`w:val="${escaped}"`);
  } finally {
    editor.destroy();
  }
  const conflict = mount(
    docxFixture(body, {
      styles:
        `<w:styles xmlns:w="${W}">${style}` +
        `<w:style w:type="character" w:styleId="${inlineCodeStyleId(base)}"><w:name w:val="Unrelated"/></w:style></w:styles>`,
    })
  );
  try {
    const before = fingerprint(conflict);
    expect(
      applyInlineCode(
        conflict,
        { paragraphId: paragraphIds(conflict)[0]!, start: 0, end: 4, code: true },
        { ...context, revision: token(conflict) }
      )
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(fingerprint(conflict)).toBe(before);
  } finally {
    conflict.destroy();
  }
});

test('suggestion mode records a formatting revision that accepts and rejects exactly', async () => {
  for (const action of ['accept', 'reject'] as const) {
    const editor = mount(
      docxFixture('<w:p><w:r><w:t xml:space="preserve">one two</w:t></w:r></w:p>')
    );
    try {
      const id = paragraphIds(editor)[0]!,
        before = fingerprint(editor);
      editor.surface!.setEditingMode('suggest');
      apply(editor, { paragraphId: id, start: 4, end: 7, code: true });
      expect(editor.surface!.editingMode()).toBe('suggest');
      expect(editor.getReviewItems()).toHaveLength(1);
      expect(await committed(editor)).toBe('one two');
      expect(await working(editor)).toBe('one `two`');
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(editor)).toBe(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      expect(editor.getReviewItems()).toEqual([]);
      expect(await working(editor)).toBe(action === 'accept' ? 'one `two`' : 'one two');
    } finally {
      editor.destroy();
    }
  }
});

test('stale reads, view mode, conflicts, and malformed requests refuse before writing', async () => {
  const editor = mount(docxFixture('<w:p><w:r><w:t>text</w:t></w:r></w:p>'));
  try {
    const id = paragraphIds(editor)[0]!,
      stale = token(editor);
    editor.surface!.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: id, offset: 4, text: '!' },
    ]);
    const before = fingerprint(editor);
    expect(
      applyInlineCode(
        editor,
        { paragraphId: id, start: 0, end: 4, code: true },
        {
          ...context,
          revision: stale,
        }
      )
    ).toMatchObject({ success: false, code: 'stale-revision' });
    for (const request of [
      { paragraphId: id, start: 2, end: 2, code: true },
      { paragraphId: id, start: 0, end: 9, code: true },
      { paragraphId: 'missing', start: 0, end: 1, code: true },
      { paragraphId: id, start: 0, end: 1, code: 'yes' },
    ] as unknown as InlineCodeRequest[])
      expect(
        applyInlineCode(editor, request, { ...context, revision: token(editor) }).success
      ).toBe(false);
    editor.surface!.setEditingMode('view');
    expect(
      applyInlineCode(
        editor,
        { paragraphId: id, start: 0, end: 4, code: true },
        {
          ...context,
          revision: token(editor),
        }
      )
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(fingerprint(editor)).toBe(before);
    expect(findNode(editor.surface!.session.part(), id)).toBeTruthy();
  } finally {
    editor.destroy();
  }
  const conflicting = mount(
    docxFixture('<w:p><w:r><w:t>text</w:t></w:r></w:p>', {
      styles: `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="SciobotCode"><w:name w:val="Other"/></w:style></w:styles>`,
    })
  );
  try {
    const pkg = conflicting.surface!.session.currentPackage();
    expect(inlineCodeStyleState(pkg)).toBe('conflict');
    expect(withInlineCodeStyle(pkg)).toBeNull();
    expect(
      inlineCodeOps(pkg, {
        paragraphId: paragraphIds(conflicting)[0]!,
        start: 0,
        end: 4,
        code: true,
      })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
  } finally {
    conflicting.destroy();
  }
});
