import { expect, test } from 'bun:test';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import {
  applicationFragmentOptions,
  applicationFragmentTarget,
} from '../../test/support/application-fragments';
import { recomposeOwnedProposals } from './recompose';
import { readReviewGroups } from './groups';

test('inserting into an unselected presence subtree refuses before publication and preserves actionable groups', async () => {
  const read = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
  if (!read.ok) throw Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const initial = await recomposeOwnedProposals(
      editor,
      {
        revision: propertyToken(editor),
        targetMarkdown: applicationFragmentTarget('First'),
        ownership: [],
      },
      propertyContext,
      {
        target: applicationFragmentOptions('first'),
        payloads: { presence: { proposalId: 'first-birth' } },
      }
    );
    if (!initial.success) throw Error(initial.error);
    const before = editor.surface!.session.currentPackage();
    const bytes = writeOoxmlPackage(before);
    const groups = readReviewGroups(before);
    expect(groups.success && groups.groups.every((group) => !group.readOnly)).toBe(true);
    let publications = 0;
    const stop = editor.surface!.session.subscribe(() => publications++);
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision: propertyToken(editor),
        targetMarkdown: applicationFragmentTarget('First').replace(
          '</phase>',
          '<activity name="Unowned" duration="2">\nNew body\n</activity>\n</phase>'
        ),
        ownership: [],
      },
      propertyContext,
      {
        target: applicationFragmentOptions('unowned'),
        payloads: { presence: { proposalId: 'unowned-birth' } },
      }
    );
    stop();
    expect(result.success).toBe(false);
    if (!result.success)
      expect(result.error).toBe('Replacement would change an unselected review group.');
    expect(publications).toBe(0);
    expect(editor.surface!.session.currentPackage()).toBe(before);
    expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(bytes);
    expect(readReviewGroups(editor.surface!.session.currentPackage())).toEqual(groups);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.surface!.session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(bytes);
  } finally {
    editor.destroy();
  }
}, 30000);

test('an outside text proposal retains unselected actionable presence groups', async () => {
  const read = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
  if (!read.ok) throw Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const first = await recomposeOwnedProposals(
      editor,
      {
        revision: propertyToken(editor),
        targetMarkdown: applicationFragmentTarget('First'),
        ownership: [],
      },
      propertyContext,
      {
        target: applicationFragmentOptions('first'),
        payloads: { presence: { proposalId: 'first-birth' } },
      }
    );
    if (!first.success) throw Error(first.error);
    const before = readReviewGroups(editor.surface!.session.currentPackage());
    if (!before.success) throw Error(before.error);
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision: propertyToken(editor),
        targetMarkdown: applicationFragmentTarget('First').replace('Outside', 'Outside updated'),
        ownership: [],
      },
      propertyContext,
      { target: applicationFragmentOptions('outside') }
    );
    if (!result.success) throw Error(result.error);
    expect(result.ownership).toHaveLength(1);
    const after = readReviewGroups(editor.surface!.session.currentPackage());
    if (!after.success) throw Error(after.error);
    for (const group of before.groups) {
      const current = after.groups.find((candidate) => candidate.id === group.id)!;
      expect(current.readOnly).toBe(false);
      expect(current.properties).toEqual(group.properties);
    }
  } finally {
    editor.destroy();
  }
});
