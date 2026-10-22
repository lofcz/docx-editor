import { expect, test } from 'bun:test';
import { readOoxmlPackage, collectRevisionSites } from '@docx-editor.dev/core/store';
import {
  cases,
  mount,
  context,
  projection,
  ast,
  fingerprint,
} from '../../test/support/link-structure-execution';
import { propertyEditor } from '../../test/support/property-execution';
import { groupToken, groups } from '../../test/support/review-groups';
import { recomposeOwnedProposals } from './recompose';
import type { OwnedProposalReplacementResult } from './replace-owned-proposals';
type Success = Extract<OwnedProposalReplacementResult, { success: true }>;
async function wave(editor: ReturnType<typeof mount>, text: string, prior?: Success) {
  let publications = 0;
  const stop = editor.surface!.session.subscribe(() => publications++);
  const result = await recomposeOwnedProposals(
    editor,
    {
      revision: groupToken(editor),
      targetMarkdown: text + '\n\nTail',
      ownership: prior?.ownership ?? [],
      contentProofs: prior?.contentProofs ?? [],
    },
    context
  );
  stop();
  if (!result.success) throw Error(`${result.code}: ${result.error}`);
  expect(result.committed).toBe(true);
  expect(publications).toBe(1);
  expect(ast((await projection(editor)).text)).toBe(ast(text + '\n\nTail'));
  return result;
}
for (const fixture of cases)
  for (const action of ['accept', 'reject'] as const)
    test(`later-wave link ${fixture.name}: ${action}`, async () => {
      const editor = mount(fixture.body);
      try {
        const original = ast((await projection(editor)).text);
        const first = await wave(editor, fixture.target),
          firstPrint = fingerprint(editor);
        const target = '[**Alpha later end**](https://example.org/later "Later")';
        const second = await wave(editor, target, first),
          secondPrint = fingerprint(editor);
        expect(
          groups(editor)
            .map((group) => group.id)
            .sort()
        ).toEqual(
          second.ownership
            .filter((group) => !first.ownership.some((old) => old.groupId === group.groupId))
            .map((group) => group.groupId)
            .sort()
        );
        expect(
          first.ownership.every((group) =>
            second.ownership.some((entry) => entry.groupId === group.groupId)
          )
        ).toBe(true);
        expect(ast((await projection(editor, true)).text)).toBe(original);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(firstPrint);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(secondPrint);
        const third = await wave(editor, 'Alpha final end', second);
        const pendingIds = groups(editor)
          .map((group) => group.id)
          .sort();
        expect(third.ownership.length).toBeGreaterThan(second.ownership.length);
        expect(ast((await projection(editor, true)).text)).toBe(original);
        const read = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!read.ok) throw Error(read.reason);
        const again = propertyEditor(read.package);
        try {
          expect(ast((await projection(again)).text)).toBe(ast('Alpha final end\n\nTail'));
          expect(
            groups(again)
              .map((group) => group.id)
              .sort()
          ).toEqual(pendingIds);
          expect(again.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          expect(ast((await projection(again)).text)).toBe(
            action === 'accept' ? ast('Alpha final end\n\nTail') : original
          );
        } finally {
          again.destroy();
        }
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(ast((await projection(editor)).text)).toBe(
          action === 'accept' ? ast('Alpha final end\n\nTail') : original
        );
        expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(0);
      } finally {
        editor.destroy();
      }
    }, 30000);
