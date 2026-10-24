import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { fontLinkDocument, selectFontLink } from '../../test/support/font-link-document';
import { applyManualLink, removeManualLink, activateManualLink, createLinkToolbar } from './links';
test('link insert/retarget/unlink retain label, selection, exact history and Word output', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  try {
    const s = editor.surface!;
    selectFontLink(editor, 0, 5);
    const before = canonicalOoxmlFingerprint(s.session.part());
    expect(applyManualLink(editor, 'https://example.com/old', 'Tip').ok).toBe(true);
    const inserted = canonicalOoxmlFingerprint(s.session.part());
    selectFontLink(editor, 2);
    const link = s.hyperlinks.linkAtCaret()!;
    expect(link.text).toBe('Alpha');
    expect(applyManualLink(editor, 'https://example.com/new').ok).toBe(true);
    const retargeted = canonicalOoxmlFingerprint(s.session.part());
    expect(s.hyperlinks.linkAtCaret()!.id).toBe(link.id);
    expect(s.hyperlinks.linkAtCaret()!.text).toBe('Alpha');
    expect(s.hyperlinks.linkAtCaret()!.tooltip).toBe('Tip');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(inserted);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(retargeted);
    const read = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!read.ok) throw new Error(read.reason);
    expect(
      canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
    ).toEqual(retargeted);
    expect(removeManualLink(editor).ok).toBe(true);
    expect(paragraphTextOf(s.session.part(), s.session.paragraphIds()[0]!)).toBe('Alpha beta');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(retargeted);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
  } finally {
    editor.destroy();
  }
});
test('unsafe targets leave every package part unchanged; internal activation uses native bookmark navigation', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  try {
    const s = editor.surface!;
    selectFontLink(editor, 0, 5);
    const before = semanticDigest(s.session.currentPackage().parts.values());
    for (const target of ['javascript:alert(1)', 'java\nscript:alert(1)', 'data:text/html,bad', ''])
      expect(applyManualLink(editor, target).ok).toBe(false);
    expect(semanticDigest(s.session.currentPackage().parts.values())).toEqual(before);
    expect(applyManualLink(editor, '#target').ok).toBe(true);
    selectFontLink(editor, 2);
    expect(activateManualLink(editor)).toBe(true);
    expect(s.state().selection.head.paragraphId).toBe(s.session.paragraphIds()[1]!);
  } finally {
    editor.destroy();
  }
});
test('opening and cancelling the native link dialog keeps the document selection', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  const toolbar = createLinkToolbar(editor);
  try {
    selectFontLink(editor, 1, 4);
    const before = editor.surface!.state().selection;
    toolbar.open();
    expect(toolbar.element.querySelector('[role="dialog"]')!.hasAttribute('hidden')).toBe(false);
    expect(editor.surface!.state().selection).toEqual(before);
    toolbar.close();
    expect(editor.surface!.state().selection).toEqual(before);
  } finally {
    toolbar.dispose();
    editor.destroy();
  }
});

test('unchanged link targets are successful no-ops and a viewing editor cannot allocate URLs', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  try {
    selectFontLink(editor, 0, 5);
    expect(applyManualLink(editor, 'https://example.com').ok).toBe(true);
    selectFontLink(editor, 2);
    const revision = editor.surface!.session.packageRevision();
    expect(applyManualLink(editor, 'https://example.com')).toEqual({ ok: true, changed: false });
    expect(editor.surface!.session.packageRevision()).toBe(revision);
  } finally {
    editor.destroy();
  }
  const viewer = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
    mode: 'view',
  });
  try {
    selectFontLink(viewer, 0, 5);
    const before = semanticDigest(viewer.surface!.session.currentPackage().parts.values());
    expect(applyManualLink(viewer, 'https://example.com').ok).toBe(false);
    expect(semanticDigest(viewer.surface!.session.currentPackage().parts.values())).toEqual(before);
  } finally {
    viewer.destroy();
  }
});
test('a link form opened for an old document cannot edit the replacement document', async () => {
  const editor = createDocxEditor({
      container: document.createElement('div'),
      document: fontLinkDocument(),
    }),
    toolbar = createLinkToolbar(editor);
  try {
    selectFontLink(editor, 0, 5);
    toolbar.open();
    await editor.load(fontLinkDocument());
    const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
    toolbar.element.querySelector('input')!.value = 'https://example.com/wrong-document';
    Array.from(toolbar.element.querySelectorAll('button'))
      .find((b) => b.textContent === 'Apply link')!
      .click();
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(before);
    expect(toolbar.element.querySelector('output')!.textContent).toContain('document changed');
  } finally {
    toolbar.dispose();
    editor.destroy();
  }
});
