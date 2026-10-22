import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  openHeadlessDocument,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import { reviewDocumentFixture } from '../../test/support/review-document';
import { docxFixture } from '../../test/support/document';
import { markdownRangeAt } from '../projection/markdown';
import { readEditorRevision } from '../projection/revision-cache';
import { findAndReplace } from '../tools/find-replace';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';

test('concurrent committed and working reads preserve pending replacements and native formatting', async () => {
  const container = document.createElement('div');
  const editor = createDocxEditor({ container, document: reviewDocumentFixture() });
  try {
    const surface = editor.surface!;
    const session = surface.session;
    const token = readEditorRevision(editor, 'doc', 'load-1')!;
    const before = canonicalOoxmlFingerprint(session.part());
    const sites = collectRevisionSites(session.part()).length;
    const displayMode = surface.layout().displayMode;
    const [committed, working] = await Promise.all([
      readCommittedView(session, token),
      readWorkingView(session, token),
    ]);
    if (!committed.success) throw new Error(committed.error);
    if (!working.success) throw new Error(working.error);
    expect(committed.view.projection.text).toBe('*Kept* Old\n\n## Heading');
    expect(working.view.projection.text).toBe('**Kept** New\n\n# Heading');
    expect(working.view.kind).toBe('working');
    expect(working.view.layout.displayMode).toBe('proposed');
    expect(working.view.projection.text.match(/New/g)).toHaveLength(1);
    expect(findAndReplace(working.view.projection.text, 'New', 'Better', false)).toEqual({
      success: true,
      newContent: '**Kept** Better\n\n# Heading',
    });
    expect(markdownRangeAt(working.view.projection, 9, 12)).toMatchObject({
      start: { paragraphId: session.paragraphIds()[0]!, offset: 8 },
      end: { paragraphId: session.paragraphIds()[0]!, offset: 11 },
    });
    expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
    expect(collectRevisionSites(session.part()).length).toBe(sites);
    expect(session.packageRevision()).toBe(token.revision);
    expect([session.canUndo(), session.canRedo()]).toEqual([false, false]);
    expect(surface.layout().displayMode).toBe(displayMode);
  } finally {
    editor.destroy();
  }
});

test('working reads preserve user history and survive save/reopen with proposals still pending', async () => {
  const container = document.createElement('div');
  const editor = createDocxEditor({ container, document: reviewDocumentFixture() });
  try {
    const surface = editor.surface!;
    const id = surface.session.paragraphIds()[1]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    surface.insertPlainText('User ');
    const read = () =>
      readWorkingView(surface.session, readEditorRevision(editor, 'doc', 'load-1')!);
    const typed = await read();
    expect(typed.success && typed.view.projection.text).toBe('**Kept** New\n\n# User Heading');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const undone = await read();
    expect(undone.success && undone.view.projection.text).toBe('**Kept** New\n\n# Heading');
    expect(surface.session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const pending = collectRevisionSites(surface.session.part()).length;
    const beforeSave = canonicalOoxmlFingerprint(surface.session.part());
    const loaded = openHeadlessDocument(new Uint8Array(await editor.save()));
    if (!loaded.ok) throw new Error(loaded.reason);
    const token = { documentId: 'doc', epoch: 'reopened', revision: loaded.view.packageRevision() };
    const [working, committed] = await Promise.all([
      readWorkingView(loaded.view, token),
      readCommittedView(loaded.view, token),
    ]);
    expect(working.success && working.view.projection.text).toBe('**Kept** New\n\n# User Heading');
    expect(committed.success && committed.view.projection.text).toBe(
      '*Kept* Old\n\n## User Heading'
    );
    expect(collectRevisionSites(loaded.view.part()).length).toBe(pending);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(beforeSave);
  } finally {
    editor.destroy();
  }
});

test('table view syntax is explicit, keeps the GFM default, and preserves native state and prior redo', async () => {
  const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const bytes = docxFixture(
    '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>' +
      p('Merged') +
      '</w:tc></w:tr><w:tr><w:tc>' +
      p('A') +
      '</w:tc><w:tc>' +
      p('B') +
      '</w:tc></w:tr></w:tbl>' +
      p('Tail')
  );
  const editor = createDocxEditor({ container: document.createElement('div'), document: bytes });
  try {
    const surface = editor.surface!,
      session = surface.session,
      tail = session.paragraphIds().at(-1)!;
    expect(
      surface.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: tail, offset: 4, text: ' manual' },
      ]).committed
    ).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const before = session.currentPackage(),
      token = readEditorRevision(editor, 'syntax', 'load-1')!,
      display = surface.layout().displayMode;
    const reads = await Promise.all([
      readWorkingView(session, token),
      readWorkingView(session, token, { tableSyntax: 'gfm' }),
      readWorkingView(session, token, { tableSyntax: 'auto' }),
      readWorkingView(session, token, { tableSyntax: 'html' }),
      readCommittedView(session, token, { tableSyntax: 'auto' }),
    ]);
    const projected = reads.map((read) => {
      if (!read.success) throw new Error(read.error);
      return read.view.projection;
    });
    expect(projected[0]!.text).toBe(projected[1]!.text);
    expect(projected[0]!.text).not.toContain('<table>');
    for (const value of projected.slice(2)) {
      expect(value.text).toContain('colspan="2"');
      expect(value.cells.map((cell) => cell.id)).toEqual(
        projected[0]!.cells.map((cell) => cell.id)
      );
    }
    expect(session.currentPackage()).toBe(before);
    expect(readEditorRevision(editor, 'syntax', 'load-1')).toEqual(token);
    expect(surface.layout().displayMode).toBe(display);
    expect(session.canRedo()).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), tail)).toBe('Tail manual');
  } finally {
    editor.destroy();
  }
}, 30_000);
