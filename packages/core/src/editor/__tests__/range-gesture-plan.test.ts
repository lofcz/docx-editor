import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor';
import { blankDocumentBytes } from '../blank-document';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '../../store/index';
test('native host range planning keeps selection and model stable and orders a reversed range', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blankDocumentBytes(),
  });
  try {
    const s = editor.surface!;
    s.insertPlainText('Alpha');
    s.flushPendingInput();
    const id = s.session.paragraphIds()[0]!,
      selection = s.state().selection,
      before = canonicalOoxmlFingerprint(s.session.part()),
      version = s.session.packageRevision();
    const plan = s.planRangeDeletion({
      anchor: { paragraphId: id, offset: 4 },
      head: { paragraphId: id, offset: 1 },
    });
    expect(plan.ops).toEqual([{ op: 'deleteText', paragraphId: id, start: 1, end: 4 }]);
    expect(plan.replaceAt).toEqual({ paragraphId: id, offset: 1 });
    expect(s.state().selection).toEqual(selection);
    expect(s.session.packageRevision()).toBe(version);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(s.applyAutomationOps(() => plan.ops)).toMatchObject({
      committed: true,
      rejected: false,
    });
    expect(paragraphTextOf(s.session.part(), id)).toBe('Aa');
  } finally {
    editor.destroy();
  }
});

test('dependent automation operations share the native transaction and roll back together', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blankDocumentBytes(),
  });
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      before = canonicalOoxmlFingerprint(s.session.part());
    const failed = s.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId: id, offset: 0, text: 'A' }],
      { kind: 'body' },
      undefined,
      false,
      true,
      'after',
      (ctx) => {
        expect(ctx.apply({ op: 'insertText', paragraphId: id, offset: 1, text: 'B' })).toBe(true);
        expect(ctx.apply({ op: 'insertText', paragraphId: 'missing', offset: 0, text: 'C' })).toBe(
          false
        );
      }
    );
    expect(failed.rejected).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(editor.exec({ type: 'undo' }).ok).toBe(false);
    expect(
      s.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: id, offset: 0, text: 'A' }],
        { kind: 'body' },
        undefined,
        false,
        true,
        'after',
        (ctx) => {
          expect(paragraphTextOf(ctx.currentPackage().parts.get(s.session.part().name)!, id)).toBe(
            'A'
          );
          expect(ctx.apply({ op: 'insertText', paragraphId: id, offset: 1, text: 'B' })).toBe(true);
        }
      ).committed
    ).toBe(true);
    expect(paragraphTextOf(s.session.part(), id)).toBe('AB');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
  } finally {
    editor.destroy();
  }
});

test('viewing and missing review support refuse before a dependent automation stage runs', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blankDocumentBytes(),
  });
  let stages = 0;
  try {
    const s = editor.surface!,
      before = canonicalOoxmlFingerprint(s.session.part());
    expect(
      s.applyAutomationOps(
        () => [],
        { kind: 'body' },
        undefined,
        true,
        true,
        'after',
        () => {
          stages++;
        }
      ).rejected
    ).toBe(true);
    expect(editor.exec({ type: 'setEditingMode', mode: 'viewing' }).ok).toBe(true);
    expect(
      s.applyAutomationOps(
        () => [],
        { kind: 'body' },
        undefined,
        false,
        true,
        'after',
        () => {
          stages++;
        }
      ).rejected
    ).toBe(true);
    expect(stages).toBe(0);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
  } finally {
    editor.destroy();
  }
});
