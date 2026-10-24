import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { propertyEditor } from '../../test/support/property-execution';
import {
  manualTrackingDocument,
  manualEdits,
  manualWave,
  manualRevisionFingerprints,
  manualParagraphFingerprints,
  manualToken,
  manualWorking,
  selectManual,
} from '../../test/support/manual-tracking';
import {
  setManualSuggestionMode,
  manualSuggestionState,
  createManualSuggestionToggle,
} from './manual-tracking';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { savedTableFingerprint } from '../../test/support/manual-table-controls';
const fingerprints = (editor: ReturnType<typeof propertyEditor>) =>
  [...editor.surface!.session.currentPackage().parts].map(([name, part]) => [
    name,
    canonicalOoxmlFingerprint(part),
  ]);
for (const decision of ['accept', 'reject'] as const)
  test(`manual insert/delete/bold keep exact authors and revision identities through AI waves, ${decision}, history and Word reopen`, async () => {
    const read = readOoxmlPackage(manualTrackingDocument());
    if (!read.ok) throw new Error(read.reason);
    const editor = propertyEditor(read.package),
      toolbar = createManualSuggestionToggle(editor);
    try {
      const first = await manualWave(editor, 'ONE');
      manualEdits(editor);
      const manual = manualRevisionFingerprints(editor),
        paragraphs = manualParagraphFingerprints(editor);
      expect(manual).toHaveLength(3);
      expect(
        editor
          .getReviewItems({ placement: false })
          .filter((item) => item.kind === 'revision' && item.author === 'User')
      ).toHaveLength(3);
      expect(manual.map((m) => m.kind).sort()).toEqual([
        'generic',
        'revisionDelete',
        'revisionInsert',
      ]);
      const before = fingerprints(editor),
        second = await manualWave(editor, 'TWO', first);
      expect(manualSuggestionState(editor)).toMatchObject({ active: true, author: 'User' });
      expect(manualRevisionFingerprints(editor)).toEqual(manual);
      expect(manualParagraphFingerprints(editor)).toEqual(paragraphs);
      expect(await manualWorking(editor)).toContain('Manual new text');
      expect(await manualWorking(editor)).toContain('**Bold** text');
      const after = fingerprints(editor);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprints(editor)).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(fingerprints(editor)).toEqual(after);
      const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
      expect(
        decide(
          editor,
          second.ownership[second.ownership.length - 1]!.groupId,
          manualToken(editor),
          { documentId: manualToken(editor).documentId, epoch: manualToken(editor).epoch }
        ).success
      ).toBe(true);
      expect(manualRevisionFingerprints(editor)).toEqual(manual);
      expect(manualParagraphFingerprints(editor)).toEqual(paragraphs);
      const saved = new Uint8Array(await editor.save()),
        loaded = readOoxmlPackage(saved);
      if (!loaded.ok) throw new Error(loaded.reason);
      const fresh = propertyEditor(loaded.package);
      try {
        expect(
          manualRevisionFingerprints(fresh).map(({ nodeId: _nodeId, ...rest }) => rest)
        ).toEqual(manual.map(({ nodeId: _nodeId, ...rest }) => rest));
        expect(await manualWorking(fresh)).toBe(await manualWorking(editor));
        expect(await savedTableFingerprint(fresh)).toEqual(
          canonicalOoxmlFingerprint(editor.surface!.session.part())
        );
      } finally {
        fresh.destroy();
      }
      expect(setManualSuggestionMode(editor, false).ok).toBe(true);
      selectManual(editor, 4, 4);
      editor.surface!.insertPlainText('!');
      editor.surface!.flushPendingInput();
      expect(manualRevisionFingerprints(editor)).toEqual(manual);
      expect(manualParagraphFingerprints(editor)).toEqual(paragraphs);
    } finally {
      toolbar.dispose();
      editor.destroy();
    }
  });
test('mode changes flush typing under its original mode and keep the current selection', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: manualTrackingDocument(),
    author: 'Alice',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  try {
    selectManual(editor, 1, 0);
    editor.surface!.insertPlainText('Before ');
    expect(setManualSuggestionMode(editor, true).ok).toBe(true);
    expect(manualRevisionFingerprints(editor)).toHaveLength(0);
    editor.surface!.insertPlainText('Tracked ');
    expect(setManualSuggestionMode(editor, false).ok).toBe(true);
    expect(manualRevisionFingerprints(editor)).toHaveLength(1);
    const selection = editor.surface!.state().selection;
    expect(setManualSuggestionMode(editor, true).ok).toBe(true);
    expect(editor.surface!.state().selection).toEqual(selection);
  } finally {
    editor.destroy();
  }
});
test('a missing author or review module refuses without deferred mode changes or package writes', () => {
  for (const review of [false, true]) {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: manualTrackingDocument(),
      modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
    });
    try {
      const before = fingerprints(editor);
      expect(setManualSuggestionMode(editor, true).ok).toBe(false);
      expect(fingerprints(editor)).toEqual(before);
      editor.setAuthor('Alice');
      expect(manualSuggestionState(editor).active).toBe(false);
    } finally {
      editor.destroy();
    }
  }
});

for (const decision of ['accept', 'reject'] as const)
  test(`manual revisions retain ownership when the AI author label matches the user during ${decision}`, async () => {
    const read = readOoxmlPackage(manualTrackingDocument());
    if (!read.ok) throw new Error(read.reason);
    const editor = propertyEditor(read.package);
    try {
      const first = await manualWave(editor, 'ONE', undefined, 'User'),
        aiIds = new Set(manualRevisionFingerprints(editor).map((site) => site.nodeId));
      manualEdits(editor);
      const manual = manualRevisionFingerprints(editor).filter((site) => !aiIds.has(site.nodeId)),
        ids = new Set(manual.map((site) => site.nodeId));
      expect(manual).toHaveLength(3);
      const second = await manualWave(editor, 'TWO', first, 'User');
      expect(manualRevisionFingerprints(editor).filter((site) => ids.has(site.nodeId))).toEqual(
        manual
      );
      const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
      expect(
        decide(editor, second.ownership.at(-1)!.groupId, manualToken(editor), manualToken(editor))
          .success
      ).toBe(true);
      expect(manualRevisionFingerprints(editor).filter((site) => ids.has(site.nodeId))).toEqual(
        manual
      );
      expect(editor.getConfiguredAuthor()).toBe('User');
    } finally {
      editor.destroy();
    }
  });
