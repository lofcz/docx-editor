import { describe, expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  openHeadlessDocument,
  paragraphTextOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  paragraphPlan,
  textPlanDocument,
  TEXT_DOCUMENT_ID,
  TEXT_EPOCH,
} from '../../test/support/text-plan';
import { createEditPlan, type CanonicalEditPlan, type CanonicalEdit } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from './committed-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTextPlan, type TrackTextContext } from './track-text';
import { readWorkingView } from './working-view';

const context: TrackTextContext = Object.freeze({
  documentId: TEXT_DOCUMENT_ID,
  epoch: TEXT_EPOCH,
  author: 'ScioBot',
  date: '2026-10-19T09:00:00Z',
});
const simpleDocument = () =>
  docxFixture(
    '<w:p><w:r><w:t>The cat sat.</w:t></w:r></w:p><w:p><w:r><w:t>Tail.</w:t></w:r></w:p>'
  );
const mount = (bytes = simpleDocument(), review = true) =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });

function apply(editor: DocxEditorInstance, target: string, paragraphIndex = 0) {
  const id = editor.surface!.session.paragraphIds()[paragraphIndex]!;
  const result = trackTextPlan(editor, paragraphPlan(editor, id, target), context);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  expect(result.receiptError).toBeUndefined();
  return result;
}

async function views(editor: DocxEditorInstance) {
  const session = editor.surface!.session;
  const revision = readEditorRevision(editor, TEXT_DOCUMENT_ID, TEXT_EPOCH)!;
  const [original, proposed] = await Promise.all([
    readCommittedView(session, revision),
    readWorkingView(session, revision),
  ]);
  if (!original.success) throw new Error(original.error);
  if (!proposed.success) throw new Error(proposed.error);
  return [original.view.projection.text, proposed.view.projection.text];
}

describe('native tracked text execution', () => {
  for (const action of ['acceptReviewItem', 'rejectReviewItem'] as const)
    test(`${action} resolves one native replacement; undo restores pending state and redo restores the decision`, async () => {
      const editor = mount();
      try {
        const session = editor.surface!.session;
        const id = session.paragraphIds()[0]!;
        const result = apply(editor, 'The car sat.');
        expect(result.proposals.map((proposal) => proposal.kind).sort()).toEqual([
          'delete',
          'insert',
        ]);
        const addresses = result.proposals.flatMap((proposal) => proposal.addresses);
        expect(new Set(addresses.map((address) => address.id)).size).toBe(2);
        expect(addresses.every((address) => address.author === context.author)).toBe(true);
        expect(addresses.every((address) => address.date === '2026-10-19T09:00:00.000Z')).toBe(
          true
        );
        expect(editor.getReviewItems()).toHaveLength(1);
        const xml = serializeOoxmlPart(session.part());
        expect(xml).toContain('<w:del');
        expect(xml).toContain('<w:ins');
        expect(await views(editor)).toEqual([
          'The cat sat\\.\n\nTail\\.',
          'The car sat\\.\n\nTail\\.',
        ]);
        const pending = canonicalOoxmlFingerprint(session.part());
        expect(editor[action](editor.getReviewItems()[0]!.key).ok).toBe(true);
        expect(paragraphTextOf(session.part(), id)).toBe(
          action === 'acceptReviewItem' ? 'The car sat.' : 'The cat sat.'
        );
        expect(collectRevisionSites(session.part())).toHaveLength(0);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
        expect(editor.getReviewItems()).toHaveLength(1);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(editor.getReviewItems()).toHaveLength(0);
        expect(paragraphTextOf(session.part(), id)).toBe(
          action === 'acceptReviewItem' ? 'The car sat.' : 'The cat sat.'
        );
      } finally {
        editor.destroy();
      }
    });

  test('pending native addresses survive save/reopen and reject returns the original', async () => {
    const editor = mount();
    let reopened: DocxEditorInstance | undefined;
    try {
      const result = apply(editor, 'The car sat.');
      const bytes = new Uint8Array(await editor.save());
      reopened = mount(bytes);
      expect(await views(reopened)).toEqual([
        'The cat sat\\.\n\nTail\\.',
        'The car sat\\.\n\nTail\\.',
      ]);
      const actual = reopened
        .surface!.session.reviewItems()
        .flatMap((item) => (item.kind === 'revision' ? item.addresses : []));
      expect(actual.map((address) => [address.id, address.author, address.date]).sort()).toEqual(
        result.proposals
          .flatMap((proposal) => proposal.addresses)
          .map((address) => [address.id, address.author, address.date])
          .sort()
      );
      expect(reopened.getReviewItems()).toHaveLength(1);
      expect(reopened.rejectReviewItem(reopened.getReviewItems()[0]!.key).ok).toBe(true);
      expect(await views(reopened)).toEqual([
        'The cat sat\\.\n\nTail\\.',
        'The cat sat\\.\n\nTail\\.',
      ]);
    } finally {
      reopened?.destroy();
      editor.destroy();
    }
  });

  test('unchanged marked runs retain their identities and unknown XML through a tracked word replacement', () => {
    const editor = mount(textPlanDocument());
    try {
      const session = editor.surface!.session;
      const id = session.paragraphIds()[0]!;
      const paragraph = findNode(session.part(), id)!;
      if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
      const runs = paragraph.children.filter((node) => node.kind === 'run');
      const keep = [runs[0]!, runs[2]!].map((node) => [node.id, canonicalOoxmlFingerprint(node)]);
      apply(editor, 'Before hippopotamus after 😀.');
      for (const [nodeId, fingerprint] of keep)
        expect(canonicalOoxmlFingerprint(findNode(session.part(), nodeId!)!)).toBe(fingerprint);
      expect(editor.acceptReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
      expect(paragraphTextOf(session.part(), id)).toBe('Before hippopotamus after 😀.');
      for (const [nodeId, fingerprint] of keep)
        expect(canonicalOoxmlFingerprint(findNode(session.part(), nodeId!)!)).toBe(fingerprint);
    } finally {
      editor.destroy();
    }
  });

  for (const target of [
    'The cats sat.',
    'The ca sat.',
    'The 🧑🏽‍🚀 sat.',
    'The c\tx\ny\u001e\u001f sat.',
    'A dog sits!',
  ])
    test(`tracked target ${JSON.stringify(target)} accepts exactly and rejects exactly`, async () => {
      const editor = mount();
      try {
        const id = editor.surface!.session.paragraphIds()[0]!;
        apply(editor, target);
        const pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('The cat sat.');
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(target);
        const saved = openHeadlessDocument(new Uint8Array(await editor.save()));
        if (!saved.ok) throw new Error(saved.reason);
        expect(canonicalOoxmlFingerprint(saved.view.part())).toBe(
          canonicalOoxmlFingerprint(editor.surface!.session.part())
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('The cat sat.');
      } finally {
        editor.destroy();
      }
    });

  test('two paragraph plans commit as one undo entry without changing editing mode', () => {
    const editor = mount();
    try {
      const ids = editor.surface!.session.paragraphIds();
      const first = paragraphPlan(editor, ids[0]!, 'The car sat.');
      const second = paragraphPlan(editor, ids[1]!, 'New tail.');
      const plan = createEditPlan(first.revision, [...first.operations, ...second.operations]);
      if (!plan.success) throw new Error(plan.error);
      expect(trackTextPlan(editor, plan.plan, context).success).toBe(true);
      expect(editor.surface!.editingMode()).toBe('edit');
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(editor.getReviewItems()).toHaveLength(0);
      expect(paragraphTextOf(editor.surface!.session.part(), ids[0]!)).toBe('The cat sat.');
      expect(paragraphTextOf(editor.surface!.session.part(), ids[1]!)).toBe('Tail.');
    } finally {
      editor.destroy();
    }
  });

  test('explicit AI attribution survives suggestion mode and own pending insertion retraction', () => {
    const editor = mount();
    try {
      expect(editor.setEditingMode('suggesting').ok).toBe(true);
      const added = apply(editor, 'The cats sat.');
      expect(
        added.proposals.every((item) =>
          item.addresses.every((address) => address.author === 'ScioBot')
        )
      ).toBe(true);
      expect(editor.surface!.editingMode()).toBe('suggest');
      const removed = apply(editor, 'The cat sat.');
      expect(removed.proposals).toEqual([]);
      expect(editor.getReviewItems()).toHaveLength(0);
      expect(editor.surface!.editingMode()).toBe('suggest');
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.getReviewItems()).toHaveLength(1);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(editor.getReviewItems()).toHaveLength(0);
    } finally {
      editor.destroy();
    }
  });

  test('replacing another author’s pending insertion retains both authors and the original view', async () => {
    const editor = mount(
      docxFixture(
        '<w:p><w:r><w:t xml:space="preserve">The </w:t></w:r><w:ins w:id="10" w:author="Alice" w:date="2026-10-19T08:00:00Z"><w:r><w:t>cat</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> sat.</w:t></w:r></w:p>'
      )
    );
    try {
      apply(editor, 'The c\tx\ny sat.');
      expect(await views(editor)).toEqual(['The  sat\\.', 'The c x  \ny sat\\.']);
      const items = editor.surface!.session.reviewItems();
      expect(items.some((item) => item.kind === 'revision' && item.author === 'Alice')).toBe(true);
      expect(items.some((item) => item.kind === 'revision' && item.author === 'ScioBot')).toBe(
        true
      );
      expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
      expect(
        paragraphTextOf(editor.surface!.session.part(), editor.surface!.session.paragraphIds()[0]!)
      ).toBe('The c\tx\ny sat.');
    } finally {
      editor.destroy();
    }
  });

  test('identical text makes no revision, history entry, or proposal', () => {
    const editor = mount();
    try {
      const before = editor.surface!.session.packageRevision();
      const result = apply(editor, 'The cat sat.');
      expect(result.proposals).toEqual([]);
      expect(result.revision.revision).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });
});

describe('atomic tracked text refusal', () => {
  function snapshot(editor: DocxEditorInstance) {
    const surface = editor.surface!,
      session = surface.session;
    return {
      fingerprint: canonicalOoxmlFingerprint(session.part()),
      revision: session.packageRevision(),
      undo: session.canUndo(),
      redo: session.canRedo(),
      selection: JSON.stringify(surface.state().selection),
      mode: surface.editingMode(),
    };
  }
  function planFor(editor: DocxEditorInstance, edits: readonly CanonicalEdit[]): CanonicalEditPlan {
    const revision = readEditorRevision(editor, TEXT_DOCUMENT_ID, TEXT_EPOCH)!;
    const result = createEditPlan(revision, edits);
    if (!result.success) throw new Error(result.error);
    return result.plan;
  }
  for (const scenario of [
    'epoch',
    'document',
    'revision',
    'before',
    'bounds',
    'overlap',
    'missing',
    'header',
    'format',
    'surrogate',
    'control',
    'grapheme',
    'date',
    'author',
    'view',
    'module',
  ])
    test(`${scenario} refuses the whole batch before changing native document state`, () => {
      const editor = mount(simpleDocument(), scenario !== 'module');
      try {
        const surface = editor.surface!,
          id = surface.session.paragraphIds()[0]!;
        surface.setSelection({
          anchor: { paragraphId: id, offset: 1 },
          head: { paragraphId: id, offset: 3 },
        });
        let plan = paragraphPlan(editor, id, 'The car sat.');
        const ctx = { ...context };
        const edit = plan.operations[0]!;
        if (edit.kind !== 'text') throw new Error('Expected text edit');
        if (scenario === 'epoch') ctx.epoch = 'other-mount';
        if (scenario === 'document') ctx.documentId = 'other-document';
        if (scenario === 'revision')
          plan = { ...plan, revision: { ...plan.revision, revision: plan.revision.revision + 1 } };
        if (scenario === 'before') plan = planFor(editor, [{ ...edit, before: 'X' }]);
        if (scenario === 'bounds')
          plan = planFor(editor, [{ ...edit, start: 100, end: 101, before: 'X' }]);
        if (scenario === 'overlap') plan = planFor(editor, [edit, { ...edit, after: 'Z' }]);
        if (scenario === 'missing') plan = planFor(editor, [{ ...edit, paragraphId: 'removed' }]);
        if (scenario === 'header')
          plan = planFor(editor, [{ ...edit, partName: '/word/header1.xml' }]);
        if (scenario === 'format')
          plan = planFor(editor, [
            edit,
            {
              kind: 'format',
              partName: edit.partName,
              target: { kind: 'paragraph', paragraphId: id },
              changes: [
                {
                  name: 'keepNext',
                  before: { present: false },
                  after: { present: true, value: true },
                },
              ],
            },
          ]);
        if (scenario === 'surrogate') plan = planFor(editor, [{ ...edit, after: '\ud800' }]);
        if (scenario === 'control') plan = planFor(editor, [{ ...edit, after: '\u0000' }]);
        if (scenario === 'date') ctx.date = '2026-02-31T09:00:00Z';
        if (scenario === 'author') ctx.author = ' ';
        if (scenario === 'view') surface.setEditingMode('view');
        if (scenario === 'grapheme') {
          surface.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: id, offset: 0, text: '😀' },
          ]);
          plan = planFor(editor, [{ ...edit, start: 1, end: 2, before: '\ude00' }]);
        }
        const before = snapshot(editor);
        const result = trackTextPlan(editor, plan, ctx);
        expect(result.success).toBe(false);
        expect(snapshot(editor)).toEqual(before);
        expect(editor.getReviewItems()).toHaveLength(0);
      } finally {
        editor.destroy();
      }
    });

  test('queued user input commits before the revision check and invalidates the AI plan', () => {
    const editor = mount();
    try {
      const surface = editor.surface!,
        id = surface.session.paragraphIds()[0]!;
      const plan = paragraphPlan(editor, id, 'The car sat.');
      surface.setSelection({
        anchor: { paragraphId: id, offset: 0 },
        head: { paragraphId: id, offset: 0 },
      });
      surface.insertPlainText('User ');
      const result = trackTextPlan(editor, plan, context);
      expect(result.success).toBe(false);
      expect(!result.success && result.code).toBe('stale-revision');
      expect(paragraphTextOf(surface.session.part(), id)).toBe('User The cat sat.');
      expect(editor.getReviewItems()).toHaveLength(0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(paragraphTextOf(surface.session.part(), id)).toBe('The cat sat.');
    } finally {
      editor.destroy();
    }
  });
});
