import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import {
  LINK_CONTEXT,
  linkDocument,
  linkPlan,
  linkTargets,
  linkViews,
} from '../../test/support/link-review';
import { applyDirectPlan } from '../ai/apply-direct';
import type { NativeEditFailure } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import type { HyperlinkFormatTarget } from '../diff/run-format';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';
import type { TrackTextContext } from './tracked-inline';

const mount = (bytes = linkDocument(), review = true): DocxEditorInstance =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
function direct(plan: CanonicalEditPlan): DocxEditorInstance {
  const editor = mount(linkDocument(), false);
  const result = applyDirectPlan(editor, plan, LINK_CONTEXT);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return editor;
}
const bodyOf = (part: OoxmlPart) =>
  serializeOoxmlPart(part)
    .replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '')
    .replace(/ xmlns:r="[^"]*"/g, '')
    .replace(/ r:id="[^"]*"/g, '');

const targets: [string, HyperlinkFormatTarget, string][] = [
  ['a URL', { url: 'https://example.org/new' }, '(https://example.org/new'],
  ['a bookmark', { anchor: 'Goal' }, '(#Goal'],
  ['a tooltip', { tooltip: 'New tip' }, 'New tip'],
  ['a URL and tooltip', { url: 'https://example.org/both', tooltip: 'Both' }, 'Both'],
];
for (const [name, target, marker] of targets)
  for (const action of ['accept', 'reject'] as const)
    test(`a tracked ${name} change ${action}s to direct native editing or the original link`, async () => {
      const editor = mount();
      try {
        const session = editor.surface!.session,
          originalPrint = canonicalOoxmlFingerprint(session.part()),
          originalLinks = linkTargets(editor);
        const [before] = await linkViews(editor);
        const plan = linkPlan(editor, target);
        const expected = direct(plan);
        const result = trackHyperlinkPlan(editor, plan, LINK_CONTEXT);
        if (!result.success) throw new Error(`${result.code}: ${result.error}`);
        expect(result.proposals).toHaveLength(1);
        expect(result.proposals[0]!.addresses.every((entry) => entry.author === 'ScioBot')).toBe(
          true
        );
        // Original keeps the old target and tooltip; Proposed reads the new one. Text is equal.
        const [original, working] = await linkViews(editor);
        expect(original).toBe(before);
        expect(working).not.toBe(before);
        if (marker.startsWith('(')) expect(working).toContain(marker);
        const plain = (text: string) => text.replace(/\]\([^)]*\)/g, ']');
        expect(plain(working!)).toBe(plain(before!));
        const pending = canonicalOoxmlFingerprint(session.part());
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(editor.getReviewItems()).toHaveLength(0);
        if (action === 'accept') {
          expect(linkTargets(editor)).toEqual(linkTargets(expected));
          expect(bodyOf(session.part())).toBe(bodyOf(expected.surface!.session.part()));
        } else {
          expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
          expect(linkTargets(editor)).toEqual(originalLinks);
        }
        // The twin link keeps the shared relationship throughout.
        expect(linkTargets(editor)[1]).toEqual(['https://example.com', null]);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
        expect(session.canUndo()).toBe(false);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
        expected.destroy();
      } finally {
        editor.destroy();
      }
    }, 30_000);

test('a pending retarget survives save/reopen with stable addresses and both decisions', async () => {
  for (const action of ['accept', 'reject'] as const) {
    const editor = mount();
    try {
      const originalLinks = linkTargets(editor);
      const plan = linkPlan(editor, targets[0]![1]);
      const result = trackHyperlinkPlan(editor, plan, LINK_CONTEXT);
      if (!result.success) throw new Error(result.error);
      const reopened = mount(new Uint8Array(await editor.save()));
      try {
        const ids = reopened
          .surface!.session.reviewItems()
          .flatMap((item) => (item.kind === 'revision' ? item.addresses : []))
          .map((entry) => entry.id)
          .sort();
        expect(ids).toEqual(result.proposals[0]!.addresses.map((entry) => entry.id).sort());
        expect(await linkViews(reopened)).toEqual(await linkViews(editor));
        expect(reopened.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(linkTargets(reopened)).toEqual(
          action === 'accept'
            ? [['https://example.org/new', 'Old tip'], originalLinks[1]]
            : originalLinks
        );
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  }
}, 60_000);

test('stale, invalid, and unsupported plans refuse atomically', () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      revision = session.packageRevision(),
      print = canonicalOoxmlFingerprint(session.part());
    const refuse = (
      plan: CanonicalEditPlan,
      code: NativeEditFailure['code'],
      context: TrackTextContext = LINK_CONTEXT
    ) => {
      const result = trackHyperlinkPlan(editor, plan, context);
      expect(result.success).toBe(false);
      if (!result.success) expect(result.code).toBe(code);
      expect(session.packageRevision()).toBe(revision);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(print);
    };
    const plan = linkPlan(editor, targets[0]![1]);
    refuse(plan, 'invalid-input', { ...LINK_CONTEXT, author: '' });
    refuse({ ...plan, revision: { ...plan.revision, revision: 999 } }, 'stale-revision');
    const edit = plan.operations[0] as Extract<
      CanonicalEditPlan['operations'][number],
      { kind: 'format' }
    >;
    const wrongOld = {
      ...plan,
      operations: [
        {
          ...edit,
          changes: [
            {
              ...edit.changes[0]!,
              before: {
                present: true,
                value: { kind: 'external', authored: 'https://other.example' },
              },
            },
          ],
        },
      ],
    } as unknown as CanonicalEditPlan;
    refuse(wrongOld, 'stale-revision');
    refuse(linkPlan(editor, { tooltip: null }), 'unsupported-operation');
    expect(trackHyperlinkPlan(editor, plan, LINK_CONTEXT).success).toBe(true);
  } finally {
    editor.destroy();
  }
});

for (const mode of ['editing', 'suggesting'] as const)
  test(`hyperlink proposals preserve ${mode} mode and change no text or marks`, async () => {
    const editor = mount();
    try {
      expect(editor.setEditingMode(mode).ok).toBe(true);
      const xml = (part: OoxmlPart) => serializeOoxmlPart(part);
      const before = xml(editor.surface!.session.part());
      const result = trackHyperlinkPlan(editor, linkPlan(editor, targets[3]![1]), LINK_CONTEXT);
      if (!result.success) throw new Error(result.error);
      expect(editor.surface!.editingMode()).toBe(mode === 'editing' ? 'edit' : 'suggest');
      // Both halves copy every run property; no format records appear.
      const after = xml(editor.surface!.session.part());
      expect(after).not.toContain('rPrChange');
      expect(after.match(/<w:b\/>/g)).toHaveLength(2 * (before.match(/<w:b\/>/g)?.length ?? 0));
    } finally {
      editor.destroy();
    }
  });
