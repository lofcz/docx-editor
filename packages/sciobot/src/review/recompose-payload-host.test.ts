import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readReviewGroups } from './groups';

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
