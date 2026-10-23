import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { atomicFixture, atomicVisual, atomicListening } from '../../test/support/atomic-execution';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import type { JsonValue, JsonObject } from '../data/json';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { groups } from '../../test/support/review-groups';

function changedSyntax(value: JsonValue): JsonValue {
  if (Array.isArray(value)) return value.map(changedSyntax);
  if (value && typeof value === 'object')
    return Object.fromEntries(
      Object.entries(value).map(([key, child]) => [
        key,
        key === 'value' && child === 'original-audio'
          ? 'second-audio'
          : key === 'value' && child === 'Good morning!'
            ? 'Second line!'
            : changedSyntax(child),
      ])
    );
  return value;
}
for (const action of ['accept', 'reject'] as const)
  test(`complete visual and arbitrary React MDX targets survive two waves, ${action}, history, and Word reopen`, async () => {
    const opaque = {
      ...atomicListening,
      mdast: {
        ...(atomicListening.mdast as JsonObject),
        name: 'NovelReactPanel',
        retainedRootData: { futureCodec: 7 },
      },
    };
    const editor = propertyEditor(atomicFixture([atomicVisual, opaque]));
    const view = async (original = false) => {
      const result = await (original ? readCommittedView : readWorkingView)(
        editor.surface!.session,
        propertyToken(editor)
      );
      if (!result.success) throw Error(result.error);
      return result.view.projection.text;
    };
    const payloads = () => {
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      return ['visual', 'audio'].map((id) => index.byId.get(id)!.payload);
    };
    try {
      const initial = await view();
      const target = (wave: string) =>
        initial
          .replace('vfs://pristine.png', `vfs://${wave}.png`)
          .replace('Old picture', `${wave} picture`)
          .replace('original-audio', `${wave}-audio`)
          .replace('Good morning!', `${wave === 'second' ? 'Second' : 'First'} line!`)
          .replace('Phase body', `Phase ${wave}`);
      const options = (wave: string) => ({
        payloads: {
          atomic: {
            proposalId: `target-${wave}`,
            fallbackTargets: ['visual', 'audio'].map((blockId) => ({
              blockId,
              before: `Old ${blockId} fallback`,
              after: `${wave} ${blockId} fallback`,
            })),
          },
        },
      });
      const first = await recomposeOwnedProposals(
        editor,
        {
          revision: propertyToken(editor),
          targetMarkdown: target('first'),
          ownership: [],
          contentProofs: [],
        },
        propertyContext,
        options('first')
      );
      if (!first.success) throw Error(first.error);
      expect(await view()).toBe(target('first'));
      expect(await view(true)).toBe(initial);
      expect(payloads()).toEqual([atomicVisual, opaque]);
      const tail = editor
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === 'Keep tail')!;
      expect(
        editor.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: tail, offset: 9, text: ' manual' },
        ]).rejected
      ).toBe(false);
      const secondTarget = target('second').replace('Keep tail', 'Keep tail manual');
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const second = await recomposeOwnedProposals(
        editor,
        {
          revision: propertyToken(editor),
          targetMarkdown: secondTarget,
          ownership: first.ownership,
          contentProofs: first.contentProofs,
        },
        propertyContext,
        options('second')
      );
      stop();
      if (!second.success) throw Error(second.error);
      expect(publications).toBe(1);
      expect(await view()).toBe(secondTarget);
      const original = initial.replace('Keep tail', 'Keep tail manual');
      expect(await view(true)).toBe(original);
      const pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(await view()).toBe(target('first').replace('Keep tail', 'Keep tail manual'));
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
      const decided = (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
        editor,
        propertyToken(editor),
        propertyContext
      );
      if (!decided.success) throw Error(decided.error);
      expect(await view()).toBe(action === 'accept' ? secondTarget : original);
      expect(groups(editor)).toHaveLength(0);
      const expected =
        action === 'accept'
          ? [
              {
                ...atomicVisual,
                attributes: {
                  ...atomicVisual.attributes,
                  src: 'vfs://second.png',
                  alt: 'second picture',
                },
              },
              { ...opaque, mdast: changedSyntax(opaque.mdast) },
            ]
          : [atomicVisual, opaque];
      expect(payloads()).toEqual(expected);
      const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw Error(saved.reason);
      const reopened = propertyEditor(saved.package);
      try {
        const current = await readWorkingView(reopened.surface!.session, propertyToken(reopened));
        if (!current.success) throw Error(current.error);
        expect(current.view.projection.text).toBe(action === 'accept' ? secondTarget : original);
        const index = buildBlockIndex(reopened.surface!.session.currentPackage());
        expect(['visual', 'audio'].map((id) => index.byId.get(id)!.payload)).toEqual(expected);
      } finally {
        reopened.destroy();
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
    } finally {
      editor.destroy();
    }
  });
