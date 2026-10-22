import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groups, groupToken } from '../../test/support/review-groups';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals, type RecomposeRequest } from './recompose';
import type { OwnedProposalReplacementResult } from './replace-owned-proposals';

const context = { ...propertyContext, date: '2026-10-21T12:00:00.000Z' };
type Editor = ReturnType<typeof propertyEditor>;
type Success = Extract<OwnedProposalReplacementResult, { success: true }>;
function fixture(body = '<w:p><w:r><w:t>Alpha end</w:t></w:r></w:p>') {
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
const pkg = (editor: Editor) => editor.surface!.session.currentPackage();
const fingerprints = (source: OoxmlPackage) =>
  [...source.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]);
function request(editor: Editor, targetMarkdown: string, previous?: Success): RecomposeRequest {
  return {
    revision: groupToken(editor),
    targetMarkdown,
    ownership: previous?.ownership ?? [],
    contentProofs: previous?.contentProofs ?? [],
  };
}
async function wave(editor: Editor, target: string, previous?: Success) {
  const result = await recomposeOwnedProposals(editor, request(editor, target, previous), context);
  expect(result.success).toBe(true);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result;
}
async function views(editor: Editor) {
  const token = groupToken(editor),
    session = editor.surface!.session;
  const original = await readCommittedView(session, token),
    proposed = await readWorkingView(session, token);
  if (!original.success || !proposed.success) throw new Error('Missing review views.');
  return [original.view.projection.text, proposed.view.projection.text];
}
function manual(editor: Editor, paragraphIndex: number, offset: number, text: string) {
  expect(
    editor.surface!.applyAutomationOps(() => [
      {
        op: 'insertText',
        paragraphId: editor.surface!.session.paragraphIds()[paragraphIndex]!,
        offset,
        text,
        bias: 'right',
      },
    ]).committed
  ).toBe(true);
}

test('three complete targets replace pending inserted text without nested revisions and preserve history and reopen', async () => {
  const editor = propertyEditor(fixture());
  try {
    const first = await wave(editor, 'Alpha XYZ end');
    const beforeSecond = fingerprints(pkg(editor));
    const second = await wave(editor, 'Alpha QRS end', first);
    expect(await views(editor)).toEqual(['Alpha end', 'Alpha QRS end']);
    expect(groups(editor)).toHaveLength(1);
    expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(1);
    const afterSecond = fingerprints(pkg(editor));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(beforeSecond);
    expect(await views(editor)).toEqual(['Alpha end', 'Alpha XYZ end']);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(afterSecond);
    const third = await wave(editor, 'Alpha TUV end', second);
    expect(await views(editor)).toEqual(['Alpha end', 'Alpha TUV end']);
    expect(groups(editor)).toHaveLength(1);
    expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(1);
    expect(third.ownership.length).toBeGreaterThan(second.ownership.length);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const fresh = propertyEditor(reopened.package);
    try {
      expect(await views(fresh)).toEqual(await views(editor));
      expect(groups(fresh).map((group) => group.id)).toEqual(
        groups(editor).map((group) => group.id)
      );
      await wave(fresh, 'Alpha JKL end', third);
      expect(await views(fresh)).toEqual(['Alpha end', 'Alpha JKL end']);
    } finally {
      fresh.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('later targets retain accepted, rejected, and manually edited content using the complete ownership ledger', async () => {
  const editor = propertyEditor(
    fixture(
      '<w:p><w:r><w:t>First end</w:t></w:r></w:p><w:p><w:r><w:t>Second end</w:t></w:r></w:p><w:p><w:r><w:t>Third end</w:t></w:r></w:p>'
    )
  );
  try {
    const first = await wave(editor, 'First AAA end\n\nSecond BBB end\n\nThird CCC end');
    const current = groups(editor);
    expect(current).toHaveLength(3);
    expect(acceptReviewGroup(editor, current[0]!.id, groupToken(editor), context).success).toBe(
      true
    );
    expect(rejectReviewGroup(editor, current[1]!.id, groupToken(editor), context).success).toBe(
      true
    );
    manual(editor, 0, 0, 'Manual ');
    const before = fingerprints(pkg(editor));
    await wave(editor, 'Manual First AAA end\n\nSecond end\n\nThird DDD end', first);
    expect(await views(editor)).toEqual([
      'Manual First AAA end\n\nSecond end\n\nThird end',
      'Manual First AAA end\n\nSecond end\n\nThird DDD end',
    ]);
    expect(groups(editor)).toHaveLength(1);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  } finally {
    editor.destroy();
  }
});

test('returned ownership recognizes an older proposal restored by undo', async () => {
  const editor = propertyEditor(fixture());
  try {
    const first = await wave(editor, 'Alpha XYZ end');
    const second = await wave(editor, 'Alpha QRS end', first);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await wave(editor, 'Alpha MNO end', second);
    expect(await views(editor)).toEqual(['Alpha end', 'Alpha MNO end']);
    expect(groups(editor)).toHaveLength(1);
  } finally {
    editor.destroy();
  }
});

test('an edit during asynchronous preparation refuses the stale target and a fresh call retains that edit', async () => {
  const editor = propertyEditor(fixture());
  try {
    const pending = recomposeOwnedProposals(editor, request(editor, 'Alpha XYZ end'), context);
    manual(editor, 0, 0, 'Manual ');
    const before = pkg(editor),
      revision = groupToken(editor);
    expect(await pending).toMatchObject({ success: false, code: 'stale-revision' });
    expect(pkg(editor)).toBe(before);
    expect(groupToken(editor)).toEqual(revision);
    expect(groups(editor)).toHaveLength(0);
    await wave(editor, 'Manual Alpha QRS end');
    expect(await views(editor)).toEqual(['Manual Alpha end', 'Manual Alpha QRS end']);
  } finally {
    editor.destroy();
  }
});

test('manual text inside an owned insertion invalidates its original content proof without writes', async () => {
  const editor = propertyEditor(fixture());
  try {
    const first = await wave(editor, 'Alpha XYZ end');
    manual(editor, 0, 8, 'MANUAL');
    const before = pkg(editor),
      revision = groupToken(editor),
      priorGroups = groups(editor);
    let publications = 0;
    const unsubscribe = editor.surface!.session.subscribe(() => publications++);
    try {
      expect(
        await recomposeOwnedProposals(editor, request(editor, 'Alpha QRS end', first), context)
      ).toMatchObject({ success: false, code: 'ambiguous-target' });
      expect(pkg(editor)).toBe(before);
      expect(groupToken(editor)).toEqual(revision);
      expect(groups(editor)).toEqual(priorGroups);
      expect(publications).toBe(0);
      expect(
        paragraphTextOf(editor.surface!.session.part(), editor.surface!.session.paragraphIds()[0]!)
      ).toContain('MANUAL');
    } finally {
      unsubscribe();
    }
  } finally {
    editor.destroy();
  }
});

test('the same complete target is a no-op and preserves an existing redo unit', async () => {
  const editor = propertyEditor(fixture());
  try {
    const first = await wave(editor, 'Alpha XYZ end');
    manual(editor, 0, 13, ' redo');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const before = pkg(editor),
      revision = groupToken(editor);
    let publications = 0;
    const unsubscribe = editor.surface!.session.subscribe(() => publications++);
    try {
      const same = await wave(editor, 'Alpha XYZ end', first);
      expect(same.committed).toBe(false);
      expect(same.ownership).toEqual(first.ownership);
      expect(pkg(editor)).toBe(before);
      expect(groupToken(editor)).toEqual(revision);
      expect(publications).toBe(0);
    } finally {
      unsubscribe();
    }
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect((await views(editor))[1]).toBe('Alpha XYZ end redo');
  } finally {
    editor.destroy();
  }
});

test('a target needing unsupported code-block creation refuses the complete wave atomically', async () => {
  const editor = propertyEditor(fixture());
  try {
    const first = await wave(editor, 'Alpha XYZ end');
    const before = pkg(editor),
      revision = groupToken(editor),
      priorGroups = groups(editor);
    let publications = 0;
    const unsubscribe = editor.surface!.session.subscribe(() => publications++);
    try {
      expect(
        (
          await recomposeOwnedProposals(
            editor,
            request(editor, 'Alpha QRS end\n\n```ts\nconst value = 1;\n```', first),
            context
          )
        ).success
      ).toBe(false);
      expect(pkg(editor)).toBe(before);
      expect(groupToken(editor)).toEqual(revision);
      expect(groups(editor)).toEqual(priorGroups);
      expect(publications).toBe(0);
    } finally {
      unsubscribe();
    }
  } finally {
    editor.destroy();
  }
});

test('later complete targets preserve an unrelated same-author pending insertion exactly', async () => {
  const editor = propertyEditor(
    fixture(
      '<w:p><w:r><w:t>Alpha end</w:t></w:r></w:p><w:p><w:r><w:t>Foreign</w:t></w:r><w:ins w:id="22" w:author="ScioBot"><w:r><w:t xml:space="preserve"> pending</w:t></w:r></w:ins></w:p>'
    )
  );
  try {
    const foreign = collectRevisionSites(editor.surface!.session.part())[0]!.node;
    const first = await wave(editor, 'Alpha XYZ end\n\nForeign pending');
    expect(first.ownership).toHaveLength(1);
    await wave(editor, 'Alpha QRS end\n\nForeign pending', first);
    expect(
      collectRevisionSites(editor.surface!.session.part()).find(
        (site) => site.node.id === foreign.id
      )!.node
    ).toBe(foreign);
    expect(groups(editor)).toHaveLength(2);
    expect(await views(editor)).toEqual([
      'Alpha end\n\nForeign',
      'Alpha QRS end\n\nForeign pending',
    ]);
  } finally {
    editor.destroy();
  }
});

test('active ownership without its original content proof refuses without removing the proposal', async () => {
  const editor = propertyEditor(fixture());
  try {
    const first = await wave(editor, 'Alpha XYZ end');
    const before = pkg(editor),
      revision = groupToken(editor);
    expect(
      (
        await recomposeOwnedProposals(
          editor,
          { revision, targetMarkdown: 'Alpha QRS end', ownership: first.ownership },
          context
        )
      ).success
    ).toBe(false);
    expect(pkg(editor)).toBe(before);
    expect(groupToken(editor)).toEqual(revision);
    expect(await views(editor)).toEqual(['Alpha end', 'Alpha XYZ end']);
  } finally {
    editor.destroy();
  }
});

test('retained host ownership and content proofs replace a saved proposal after reload without adopting foreign groups', async () => {
  const editor = propertyEditor(
    fixture(
      '<w:p><w:r><w:t>Alpha end</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>Foreign</w:t></w:r>' +
        '<w:ins w:id="22" w:author="ScioBot"><w:r><w:t xml:space="preserve"> pending</w:t></w:r></w:ins></w:p>'
    )
  );
  try {
    const first = await wave(editor, 'Alpha XYZ end\n\nForeign pending');
    // The host retains these plain receipts independently from DOCX and live canonical node IDs.
    const retained = JSON.parse(
      JSON.stringify({ ownership: first.ownership, contentProofs: first.contentProofs })
    ) as Pick<Success, 'ownership' | 'contentProofs'>;
    expect(retained.ownership).toHaveLength(1);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    const fresh = propertyEditor(saved.package);
    try {
      const foreignGroup = groups(fresh).find((group) =>
        group.native.some((member) => member.address.id === '22')
      )!;
      const foreign = collectRevisionSites(fresh.surface!.session.part()).find((site) =>
        site.node.attributes.some(
          (attribute) => attribute.localName === 'id' && attribute.value === '22'
        )
      )!.node;
      const result = await recomposeOwnedProposals(
        fresh,
        {
          revision: groupToken(fresh),
          targetMarkdown: 'Alpha QRS end\n\nForeign pending',
          ...retained,
        },
        context
      );
      expect(result.success).toBe(true);
      if (!result.success) throw new Error(result.error);
      expect(result.ownership.some((owned) => owned.groupId === foreignGroup.id)).toBe(false);
      expect(groups(fresh)).toHaveLength(2);
      expect(groups(fresh).some((group) => group.id === first.ownership[0]!.groupId)).toBe(false);
      expect(
        collectRevisionSites(fresh.surface!.session.part()).find(
          (site) => site.node.id === foreign.id
        )!.node
      ).toBe(foreign);
      expect(await views(fresh)).toEqual([
        'Alpha end\n\nForeign',
        'Alpha QRS end\n\nForeign pending',
      ]);
      expect(await views(editor)).toEqual([
        'Alpha end\n\nForeign',
        'Alpha XYZ end\n\nForeign pending',
      ]);
    } finally {
      fresh.destroy();
    }
  } finally {
    editor.destroy();
  }
});
