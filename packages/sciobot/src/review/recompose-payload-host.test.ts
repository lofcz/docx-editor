import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  writeOoxmlPackageKeepingRevisionIds,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { propertyRecordFixture } from '../../test/support/property-record';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readReviewGroups } from './groups';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

for (const invalidation of ['cancel', 'owner'] as const)
  test(`host ${invalidation} during deferred payload preparation publishes no shared proposal`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackageKeepingRevisionIds(propertyRecordFixture()),
      documentId: propertyContext.documentId,
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    let resume!: () => void;
    const pending = new Promise<void>((resolve) => {
      resume = resolve;
    });
    try {
      const [editor] = pair.editors;
      const surface = editor!.surface!;
      const savedBefore = await pair.savedDigests();
      const before = pair.editors.map((peer) => peer.surface!.session.currentPackage());
      const revision = propertyToken(editor!);
      const working = await readWorkingView(surface.session, revision);
      if (!working.success) throw Error(working.error);
      const request = {
        revision,
        targetMarkdown: working.view.projection.text
          .replace('name="Old activity"', 'name="Host change"')
          .replace('Body', 'Body changed'),
        ownership: [],
      };
      const controller = new AbortController();
      const capturedOwner = {};
      let owner = capturedOwner;
      let entered!: () => void;
      const preparing = new Promise<void>((resolve) => {
        entered = resolve;
      });
      const payloads = {
        properties: {
          proposalId: 'guarded-host-property',
          fallbackTargets: [
            {
              blockId: 'activity',
              before: 'Old activity — 10 minutes',
              after: 'Host change — 10 minutes',
            },
          ],
        },
      };
      const result = recomposeOwnedProposals(editor!, request, propertyContext, {
        isCurrent: () => !controller.signal.aborted && owner === capturedOwner,
        preparePayloads: async () => {
          entered();
          await pending;
          return payloads;
        },
      });
      await preparing;
      if (invalidation === 'cancel') controller.abort();
      else owner = {};
      // Neither cancellation nor a host-owner change advances the native revision.
      expect(propertyToken(editor!)).toEqual(revision);
      resume();
      expect(await result).toMatchObject({ success: false, code: 'document-not-ready' });
      await pair.settle();
      for (const [index, peer] of pair.editors.entries()) {
        expect(peer.surface!.session.currentPackage()).toBe(before[index]);
        expect(peer.surface!.session.canUndo()).toBe(false);
        expect(readReviewGroups(peer.surface!.session.currentPackage())).toMatchObject({
          success: true,
          groups: [],
        });
      }
      expect(await pair.savedDigests()).toEqual(savedBefore);

      // A later owned request can use the unchanged document and its ordinary shared history.
      expect(
        await recomposeOwnedProposals(editor!, request, propertyContext, {
          isCurrent: () => true,
          payloads,
        })
      ).toMatchObject({ success: true, committed: true });
      await pair.settle();
      const proposed = await pair.savedDigests();
      expect(proposed[0]).toEqual(proposed[1]);
      expect(editor!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(savedBefore);
      expect(editor!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(proposed);
    } finally {
      resume();
      pair.destroy();
    }
  }, 30_000);

for (const failure of ['throw', 'bad-label', 'accessor', 'stale'] as const)
  test(`host payload ${failure} cannot partially commit a property or body suggestion`, async () => {
    const editor = propertyEditor();
    try {
      const surface = editor.surface!;
      const revision = propertyToken(editor);
      const working = await readWorkingView(surface.session, revision);
      if (!working.success) throw Error(working.error);
      const before = surface.session.currentPackage();
      const native = canonicalOoxmlFingerprint(surface.session.part());
      const result = await recomposeOwnedProposals(
        editor,
        {
          revision,
          targetMarkdown: working.view.projection.text
            .replace('name="Old activity"', 'name="Host change"')
            .replace('Body', 'Body changed'),
          ownership: [],
        },
        propertyContext,
        {
          preparePayloads: async (targets) => {
            expect(targets[0]!.beforeText).toBe('Old activity — 10 minutes');
            // Detached callback data cannot modify the canonical baseline.
            Object.assign(targets[0]!.before.attributes, { name: 'Attempted mutation' });
            if (failure === 'throw') throw Error('Host labels unavailable.');
            if (failure === 'accessor')
              return Object.defineProperty({}, 'properties', {
                enumerable: true,
                get() {
                  throw Error('Getter must not run.');
                },
              });
            if (failure === 'stale') {
              const paragraphId = surface.session
                .paragraphIds()
                .find((id) => paragraphTextOf(surface.session.part(), id) === 'Outside')!;
              expect(
                surface.applyAutomationOps(() => [
                  { op: 'insertText', paragraphId, offset: 7, text: ' peer' },
                ])
              ).toMatchObject({ committed: true, rejected: false });
            }
            return {
              properties: {
                proposalId: 'host-property',
                fallbackTargets: [
                  {
                    blockId: 'activity',
                    before: failure === 'bad-label' ? 'Wrong label' : 'Old activity — 10 minutes',
                    after: 'Host change — 10 minutes',
                  },
                ],
              },
            };
          },
        }
      );
      expect(result.success).toBe(false);
      const groups = readReviewGroups(surface.session.currentPackage());
      if (!groups.success) throw Error(groups.error);
      expect(groups.groups).toHaveLength(0);
      if (failure === 'stale') {
        expect(result).toMatchObject({ code: 'stale-revision' });
        expect(
          surface.session.paragraphIds().map((id) => paragraphTextOf(surface.session.part(), id))
        ).toContain('Outside peer');
      } else {
        expect(surface.session.currentPackage()).toBe(before);
        expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(native);
        expect(surface.session.canUndo()).toBe(false);
      }
    } finally {
      editor.destroy();
    }
  });

test('explicit property fallback options take precedence over the optional host callback', async () => {
  const editor = propertyEditor();
  try {
    const revision = propertyToken(editor);
    const working = await readWorkingView(editor.surface!.session, revision);
    if (!working.success) throw Error(working.error);
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision,
        targetMarkdown: working.view.projection.text.replace(
          'name="Old activity"',
          'name="Explicit"'
        ),
        ownership: [],
      },
      propertyContext,
      {
        preparePayloads: () => {
          throw Error('The explicit executor must bypass host preparation.');
        },
        payloads: {
          properties: {
            proposalId: 'explicit-host-precedence',
            fallbackTargets: [
              {
                blockId: 'activity',
                before: 'Old activity — 10 minutes',
                after: 'Explicit — 10 minutes',
              },
            ],
          },
        },
      }
    );
    expect(result.success).toBe(true);
    const next = await readWorkingView(editor.surface!.session, propertyToken(editor));
    if (!next.success) throw Error(next.error);
    expect(next.view.projection.text).toContain('name="Explicit"');
  } finally {
    editor.destroy();
  }
});
