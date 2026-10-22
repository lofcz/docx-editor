import { groups, groupToken, proposeGroupText } from '../../test/support/review-groups';
import { acceptReviewGroup } from './accept-group';
import { expect, test } from 'bun:test';
import {
  readOoxmlPackage,
  findNode,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import {
  cases,
  mount,
  context,
  plan,
  projection,
  ast,
  fingerprint,
} from '../../test/support/link-structure-execution';
import { propertyEditor } from '../../test/support/property-execution';
import { applyDirectPlan } from '../ai/apply-direct';
import { trackHyperlinkPlan } from './track-hyperlinks';
for (const fixture of cases)
  for (const action of ['accept', 'reject'] as const)
    test(`D27 tracked ${fixture.name}: ${action}`, async () => {
      const editor = mount(fixture.body),
        direct = mount(fixture.body);
      try {
        const original = ast((await projection(editor)).text),
          before = fingerprint(editor),
          target = fixture.target + '\n\nTail';
        const wanted = await plan(direct, target);
        const applied = applyDirectPlan(direct, wanted, context);
        if (!applied.success) throw Error(applied.error);
        const expected = fingerprint(direct);
        let publications = 0;
        const stop = editor.surface!.session.subscribe(() => publications++);
        const result = trackHyperlinkPlan(editor, await plan(editor, target), context);
        stop();
        if (!result.success) throw Error(`${result.code}: ${result.error}`);
        expect(publications).toBe(1);
        expect(result.proposals.length).toBeGreaterThan(0);
        expect(result.proposals.every((item) => item.kind === 'hyperlinkStructure')).toBe(true);
        expect(ast((await projection(editor, true)).text)).toBe(original);
        expect(ast((await projection(editor)).text)).toBe(ast(target));
        const pending = fingerprint(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(pending);
        const read = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!read.ok) throw Error(read.reason);
        const saved = propertyEditor(read.package);
        try {
          expect(ast((await projection(saved)).text)).toBe(ast(target));
          expect(saved.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          expect(fingerprint(saved)).toBe(action === 'accept' ? expected : before);
          expect(ast((await projection(saved)).text)).toBe(
            action === 'accept' ? ast(target) : original
          );
        } finally {
          saved.destroy();
        }
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(ast((await projection(editor)).text)).toBe(
          action === 'accept' ? ast(target) : original
        );
        expect(fingerprint(editor)).toBe(action === 'accept' ? expected : before);
        expect(editor.getReviewItems()).toHaveLength(0);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(pending);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      } finally {
        editor.destroy();
        direct.destroy();
      }
    }, 30000);

test('a late refused link preserves the complete package and prior redo', async () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      tail = session.paragraphIds()[1]!;
    expect(
      editor.surface!.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' manual' }],
        { kind: 'body' },
        [],
        false,
        true
      )
    ).toMatchObject({ committed: true, rejected: false });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(session.canRedo()).toBe(true);
    const before = session.currentPackage();
    const prepared = await plan(
      editor,
      '[Alpha](https://example.org/a) [end](https://example.org/b)\n\nTail'
    );
    let count = 0;
    const unsafe = {
      ...prepared,
      operations: prepared.operations.map((edit) =>
        edit.kind === 'link' && ++count === 2
          ? { ...edit, target: { kind: 'external' as const, url: 'javascript:alert(1)' } }
          : edit
      ),
    };
    let publications = 0;
    const stop = session.subscribe(() => publications++);
    const result = trackHyperlinkPlan(editor, unsafe, context);
    stop();
    expect(result.success).toBe(false);
    expect(session.currentPackage()).toBe(before);
    expect(publications).toBe(0);
    expect(session.canRedo()).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), tail)).toBe('Tail manual');
  } finally {
    editor.destroy();
  }
});

test('a link decision retains an unselected same-author proposal and its native sites', async () => {
  const editor = mount();
  try {
    proposeGroupText(editor, 'Tail', 'Tail extra');
    const originalGroups = groups(editor),
      pkg = editor.surface!.session.currentPackage();
    const sites = originalGroups.flatMap((group) =>
      group.native.flatMap((member) =>
        member.siteNodeIds.map((id) => ({
          id,
          part: member.partName,
          print: canonicalOoxmlFingerprint(findNode(pkg.parts.get(member.partName)!, id)!),
        }))
      )
    );
    const result = trackHyperlinkPlan(
      editor,
      await plan(editor, '[Alpha end](https://example.org/new)\n\nTail extra'),
      context
    );
    if (!result.success) throw Error(result.error);
    expect(result.proposals).toHaveLength(1);
    const accepted = acceptReviewGroup(
      editor,
      result.proposals[0]!.identity,
      groupToken(editor),
      context
    );
    if (!accepted.success) throw Error(accepted.error);
    expect(groups(editor).map((group) => group.id)).toEqual(
      originalGroups.map((group) => group.id)
    );
    for (const site of sites)
      expect(
        canonicalOoxmlFingerprint(
          findNode(editor.surface!.session.currentPackage().parts.get(site.part)!, site.id)!
        )
      ).toBe(site.print);
    expect(ast((await projection(editor, true)).text)).toBe(
      ast('[Alpha end](https://example.org/new)\n\nTail')
    );
    expect(ast((await projection(editor)).text)).toBe(
      ast('[Alpha end](https://example.org/new)\n\nTail extra')
    );
  } finally {
    editor.destroy();
  }
});
