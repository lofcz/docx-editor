import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { paragraphFragmentsOf } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  stylesPartOf,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import {
  taskCompletionStyleSource,
  withTaskCompletionStyles,
} from '../native/task-completion-style';
import { setManualList, setManualListChecked, manualListState } from './lists';
import { handleManualListKey, installListEditing } from './list-input';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const context = { documentId: 'completed-tasks', epoch: 'mount' };
const styleId = (editor: ReturnType<typeof createDocxEditor>, id: string) =>
  directParagraphProperties(editor.surface!.session.part(), id).find(
    (p) => p.localName === 'pStyle'
  )?.attributes?.val;
const content =
  '<w:r><w:t>Plain</w:t></w:r><w:r><w:rPr><w:color w:val="CC0000"/><w:strike w:val="0"/><w:b/></w:rPr><w:t>Styled</w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t>Struck</w:t></w:r>';
function open(base?: string, collision = '') {
  const styles = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:color w:val="112233"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Imported"><w:basedOn w:val="Normal"/><w:rPr><w:i/></w:rPr></w:style>${collision}</w:styles>`;
  const prepared = prepareNativeDocument(
    docxFixture(
      `<w:p>${base ? `<w:pPr><w:pStyle w:val="${base}"/></w:pPr>` : ''}${content}</w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`,
      { styles }
    )
  );
  if (!prepared.success) throw new Error(prepared.error);
  return createDocxEditor({
    container: document.createElement('div'),
    document: prepared.document,
    author: 'Alice',
  });
}
const spans = (editor: ReturnType<typeof createDocxEditor>, id: string) =>
  editor
    .surface!.layout()
    .pages.flatMap((page) => paragraphFragmentsOf(page))
    .filter((p) => p.paragraphId === id)
    .flatMap((p) => p.lines.flatMap((l) => l.spans));
for (const base of [undefined, 'Normal', 'Imported'])
  test(`completion preserves mixed runs and restores source style ${base ?? 'absent'}`, async () => {
    const editor = open(base);
    try {
      const surface = editor.surface!,
        id = surface.session.paragraphIds()[0]!;
      const runs = findNode(surface.session.part(), id)!;
      if (runs.kind !== 'paragraph') throw new Error('Missing paragraph');
      const before = runs.children.filter((n) => n.kind === 'run').map(canonicalOoxmlFingerprint);
      expect(setManualList(editor, 'todo', context).success).toBe(true);
      expect(setManualListChecked(editor, true, context).success).toBe(true);
      expect(
        taskCompletionStyleSource(surface.session.currentPackage(), styleId(editor, id))
      ).toEqual({ base });
      const rendered = spans(editor, id);
      expect(rendered.find((s) => s.text === 'Plain')!.style).toMatchObject({
        strike: true,
        color: '606060',
      });
      expect(rendered.find((s) => s.text === 'Styled')!.style).toMatchObject({
        strike: false,
        color: 'CC0000',
        bold: true,
      });
      expect(rendered.find((s) => s.text === 'Struck')!.style.strike).toBe(true);
      expect(rendered.find((s) => s.text === '☒')!.style).toMatchObject({
        fontSizePt: 15,
        characterSpacingPt: 9,
        strike: false,
      });
      const projected = await readWorkingView(
        surface.session,
        readEditorRevision(editor, context.documentId, context.epoch)!
      );
      if (!projected.success) throw new Error(projected.error);
      expect(projected.view.projection.text).toContain('- [x] ');
      expect(projected.view.projection.text).toContain('Plain');
      expect(projected.view.projection.text).toContain('~~Struck~~');
      expect(projected.view.projection.text).not.toContain('~~Plain');
      const paragraph = findNode(surface.session.part(), id)!;
      if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
      expect(
        paragraph.children.filter((n) => n.kind === 'run').map(canonicalOoxmlFingerprint)
      ).toEqual(before);
      const read = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!read.ok) throw new Error(read.reason);
      // Reopened tree identities differ. The document retains the owned style definition.
      expect(
        stylesPartOf(read.package)!.root.children.some(
          (n) =>
            n.kind !== 'textValue' &&
            n.attributes.some((a) => a.localName === 'styleId' && a.value === styleId(editor, id))
        )
      ).toBe(true);
      expect(setManualListChecked(editor, false, context).success).toBe(true);
      expect(styleId(editor, id)).toBe(base);
      expect(spans(editor, id).find((s) => s.text === 'Plain')!.style).toMatchObject({
        strike: false,
        color: '112233',
      });
      expect(spans(editor, id).find((s) => s.text === 'Struck')!.style.strike).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(manualListState(editor, context)!.checkbox!.checked).toBe(true);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(styleId(editor, id)).toBe(base);
    } finally {
      editor.destroy();
    }
  });

test('completion retains a source paragraph strike through the native style cascade and reopen', async () => {
  const editor = open(
    'AlreadyStruck',
    '<w:style w:type="paragraph" w:styleId="AlreadyStruck"><w:basedOn w:val="Normal"/><w:rPr><w:strike/></w:rPr></w:style>'
  );
  try {
    const id = editor.surface!.session.paragraphIds()[0]!;
    expect(spans(editor, id).find((span) => span.text === 'Plain')!.style.strike).toBe(true);
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    expect(setManualListChecked(editor, true, context).success).toBe(true);
    expect(spans(editor, id).find((span) => span.text === 'Plain')!.style.strike).toBe(true);
    const prepared = prepareNativeDocument(new Uint8Array(await editor.save()));
    if (!prepared.success) throw new Error(prepared.error);
    await editor.load(prepared.document);
    const reopenedId = editor.surface!.session.paragraphIds()[0]!;
    expect(
      taskCompletionStyleSource(
        editor.surface!.session.currentPackage(),
        styleId(editor, reopenedId)
      )
    ).toEqual({ base: 'AlreadyStruck' });
    expect(spans(editor, reopenedId).find((span) => span.text === 'Plain')!.style.strike).toBe(
      true
    );
    expect(setManualListChecked(editor, false, context).success).toBe(true);
    expect(styleId(editor, reopenedId)).toBe('AlreadyStruck');
    expect(spans(editor, reopenedId).find((span) => span.text === 'Plain')!.style.strike).toBe(
      true
    );
  } finally {
    editor.destroy();
  }
});
for (const target of ['disc', 'none'] as const)
  test(`leaving completed tasks for ${target} restores the original paragraph style`, () => {
    const editor = open('Imported');
    try {
      const id = editor.surface!.session.paragraphIds()[0]!;
      expect(setManualList(editor, 'todo', context).success).toBe(true);
      expect(setManualListChecked(editor, true, context).success).toBe(true);
      expect(setManualList(editor, target, context).success).toBe(true);
      expect(styleId(editor, id)).toBe('Imported');
      expect(manualListState(editor, context)!.checkbox).toBeNull();
    } finally {
      editor.destroy();
    }
  });
test('Enter after a completed task creates an unchecked successor with its source style', () => {
  const editor = open('Imported');
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!;
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    expect(setManualListChecked(editor, true, context).success).toBe(true);
    const p = { paragraphId: id, offset: paragraphTextOf(s.session.part(), id)!.length };
    s.setSelection({ anchor: p, head: p });
    expect(handleManualListKey(editor, 'Enter', context)).toBe(true);
    expect(manualListState(editor, context)!.checkbox!.checked).toBe(false);
    expect(styleId(editor, s.state().selection.head.paragraphId)).toBe('Imported');
  } finally {
    editor.destroy();
  }
});
test('imported style collisions remain intact and use another owned identity', () => {
  const editor = open(
    undefined,
    '<w:style w:type="paragraph" w:styleId="SciobotTaskDone__A"><w:name w:val="Foreign"/></w:style>'
  );
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!;
    const before = canonicalOoxmlFingerprint(
      stylesPartOf(s.session.currentPackage())!.root.children[2]!
    );
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    expect(setManualListChecked(editor, true, context).success).toBe(true);
    expect(styleId(editor, id)).toBe('SciobotTaskDone__A_1');
    expect(
      canonicalOoxmlFingerprint(stylesPartOf(s.session.currentPackage())!.root.children[2]!)
    ).toBe(before);
  } finally {
    editor.destroy();
  }
});

test('owned checkbox activation keeps the caret and supports keyboard focus', async () => {
  const editor = open('Imported');
  const host = editor.surface!.hostElement();
  document.body.append(host);
  const errors: string[] = [];
  const stop = installListEditing(editor, context, {
    onError: (message) => errors.push(message),
    label: () => 'Task',
  });
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      tail = s.session.paragraphIds()[1]!;
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    const point = { paragraphId: tail, offset: 2 };
    s.setSelection({ anchor: point, head: point });
    await Promise.resolve();
    const widget = () => host.querySelector<HTMLElement>('[data-docx-cc-widget="checkbox"]')!;
    expect(widget().tabIndex).toBe(0);
    expect(widget().getAttribute('aria-label')).toBe('PlainStyledStruck');
    widget().dispatchEvent(
      new PointerEvent('pointerdown', {
        bubbles: true,
        cancelable: true,
        button: 0,
        pointerId: 1,
        pointerType: 'mouse',
      })
    );
    expect(manualListState(editor, context, id)!.checkbox!.checked).toBe(true);
    expect(taskCompletionStyleSource(s.session.currentPackage(), styleId(editor, id))).toEqual({
      base: 'Imported',
    });
    expect(s.state().selection.head).toEqual(point);
    widget().focus();
    widget().dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: ' ' })
    );
    await Promise.resolve();
    expect(manualListState(editor, context, id)!.checkbox!.checked).toBe(false);
    expect(styleId(editor, id)).toBe('Imported');
    expect(document.activeElement).toBe(widget());
    const end = { paragraphId: id, offset: paragraphTextOf(s.session.part(), id)!.length };
    s.setSelection({ anchor: end, head: end });
    s.insertPlainText(' changed');
    s.flushPendingInput();
    expect(widget().getAttribute('aria-label')).toBe('PlainStyledStruck changed');
    widget().click();
    await Promise.resolve();
    expect(manualListState(editor, context, id)!.checkbox!.checked).toBe(true);
    expect(errors).toEqual([]);
  } finally {
    stop();
    editor.destroy();
    host.remove();
  }
});

test('preparation restores legacy completed task appearance before seeding and stays idempotent', async () => {
  const editor = open('Imported');
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!;
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    expect(setManualListChecked(editor, true, context).success).toBe(true);
    expect(
      s.applyAutomationOps(() => [
        {
          op: 'setParagraphProperties',
          paragraphId: id,
          properties: directParagraphProperties(s.session.part(), id).map((p) =>
            p.localName === 'pStyle' ? { ...p, attributes: { val: 'Imported' } } : p
          ),
        },
      ]).committed
    ).toBe(true);
    expect(spans(editor, id).find((span) => span.text === 'Plain')!.style.strike).toBe(false);
    const prepared = prepareNativeDocument(new Uint8Array(await editor.save()));
    if (!prepared.success) throw new Error(prepared.error);
    expect(prepared.changed).toBe(true);
    await editor.load(prepared.document);
    const reopenedId = editor.surface!.session.paragraphIds()[0]!;
    expect(styleId(editor, reopenedId)).toBe('SciobotTaskDone__SImported');
    expect(spans(editor, reopenedId).find((span) => span.text === 'Plain')!.style.strike).toBe(
      true
    );
    const repeated = prepareNativeDocument(prepared.document);
    expect(repeated).toMatchObject({ success: true, changed: false });
    if (repeated.success) expect(repeated.document).toBe(prepared.document);
  } finally {
    editor.destroy();
  }
});

test('task listeners follow document reloads and keep read-only reasons visible', async () => {
  const editor = open();
  const host = editor.surface!.hostElement();
  document.body.append(host);
  const stop = installListEditing(editor, context, {
    label: () => 'Task',
    disabledReason: (reason) => (reason === 'readOnly' ? 'Viewing this document' : undefined),
  });
  try {
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    const document = new Uint8Array(await editor.save());
    await editor.load(document);
    const surface = editor.surface!;
    const widget = () =>
      surface.hostElement().querySelector<HTMLElement>('[data-docx-cc-widget="checkbox"]')!;
    expect(widget().tabIndex).toBe(0);
    widget().focus();
    widget().dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Enter' })
    );
    await Promise.resolve();
    expect(manualListState(editor, context)!.checkbox!.checked).toBe(true);
    expect(editor.exec({ type: 'setEditingMode', mode: 'viewing' }).ok).toBe(true);
    expect(widget().getAttribute('aria-disabled')).toBe('true');
    expect(widget().getAttribute('title')).toBe('Viewing this document');
    widget().dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: ' ' })
    );
    expect(manualListState(editor, context)!.checkbox!.checked).toBe(true);
    expect(editor.exec({ type: 'setEditingMode', mode: 'editing' }).ok).toBe(true);
    stop();
    await editor.load(document);
    expect(widget().tabIndex).toBe(-1);
  } finally {
    stop();
    editor.destroy();
    host.remove();
  }
});

test('many imported paragraph styles prepare together and remain unchanged on a second pass', () => {
  const styles = Array.from(
    { length: 1000 },
    (_, i) =>
      `<w:style w:type="paragraph" w:styleId="Imported${i}"><w:name w:val="Style ${i}"/></w:style>`
  ).join('');
  const loaded = readOoxmlPackage(
    docxFixture('<w:p/>', { styles: `<w:styles xmlns:w="${W}">${styles}</w:styles>` })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const prepared = withTaskCompletionStyles(loaded.package);
  expect(stylesPartOf(prepared)!.root.children.length).toBe(2001);
  expect(taskCompletionStyleSource(prepared, 'SciobotTaskDone__SImported999')).toEqual({
    base: 'Imported999',
  });
  expect(withTaskCompletionStyles(prepared)).toBe(prepared);
});
