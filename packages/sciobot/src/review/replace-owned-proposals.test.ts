import { expect, spyOn, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  propertyContext,
  propertyEditor,
  executeProperties,
} from '../../test/support/property-execution';
import { groupFixture, groups, groupToken, groupParagraph } from '../../test/support/review-groups';
import { docxFixture } from '../../test/support/document';
import { createEditPlan } from '../diff/plan';
import { planParagraphText } from '../diff/paragraph-text';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { readReviewGroupStore } from './group-record-store';
import type { NativeReviewGroup } from './groups';
import type { OwnedProposalGroup } from './owned-proposal-plan';
import { replaceOwnedProposals } from './replace-owned-proposals';

const context = { ...propertyContext, date: '2026-10-21T10:00:00.000Z' };
function fixture(foreign = true) {
  const read = readOoxmlPackage(
    docxFixture(
      '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' +
        '<w:ins w:id="0" w:author="ScioBot"><w:r><w:t>old</w:t></w:r></w:ins>' +
        '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>' +
        (foreign
          ? '<w:p><w:r><w:t>Foreign</w:t></w:r><w:ins w:id="22" w:author="ScioBot"><w:r><w:t xml:space="preserve"> pending</w:t></w:r></w:ins></w:p>'
          : '')
    )
  );
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
const pkg = (editor: ReturnType<typeof propertyEditor>) => editor.surface!.session.currentPackage();
const grant = (group: NativeReviewGroup): OwnedProposalGroup => ({
  groupId: group.id,
  native: group.native.map((member) => ({
    partName: member.partName,
    addresses: [member.address],
  })),
  properties: group.properties.map((member) => ({ ...member })),
});
function owned(editor: ReturnType<typeof propertyEditor>) {
  return [grant(groups(editor).find((g) => g.native.some((n) => n.address.id === '0'))!)];
}
function plan(editor: ReturnType<typeof propertyEditor>, text = 'new') {
  const session = editor.surface!.session;
  const result = createEditPlan(groupToken(editor), [
    {
      kind: 'text',
      partName: session.part().name,
      paragraphId: session.paragraphIds()[0]!,
      start: 6,
      end: 6,
      before: '',
      after: text,
    },
  ]);
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
async function views(editor: ReturnType<typeof propertyEditor>) {
  const token = groupToken(editor),
    session = editor.surface!.session;
  const [before, after] = await Promise.all([
    readCommittedView(session, token),
    readWorkingView(session, token),
  ]);
  if (!before.success || !after.success) throw new Error('Cannot read replacement views.');
  return [before.view.projection.text, after.view.projection.text];
}
const fingerprints = (source: OoxmlPackage) =>
  [...source.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]);

test('replacing a pending inserted word removes its old receipt and preserves same-author foreign markup through history and reopen', async () => {
  const editor = propertyEditor(fixture());
  try {
    const session = editor.surface!.session,
      before = fingerprints(pkg(editor)),
      selection = owned(editor);
    const foreign = collectRevisionSites(session.part()).find((s) =>
      s.node.attributes.some((a) => a.localName === 'id' && a.value === '22')
    )!;
    let publications = 0;
    const unsubscribe = session.subscribe(() => publications++);
    const result = replaceOwnedProposals(editor, selection, plan(editor), context);
    unsubscribe();
    expect(result).toMatchObject({ success: true, committed: true });
    if (!result.success) throw new Error(result.error);
    expect(publications).toBe(1);
    expect(result.ownership).toHaveLength(1);
    expect(result.ownership[0]!.native.flatMap((n) => n.addresses).every((a) => a.id !== '0')).toBe(
      true
    );
    expect(Object.isFrozen(result.ownership)).toBe(true);
    const sites = collectRevisionSites(session.part());
    expect(sites).toHaveLength(2);
    expect(sites.find((s) => s.node.id === foreign.node.id)!.node).toBe(foreign.node);
    expect(await views(editor)).toEqual([
      'Alpha  end\n\nForeign',
      'Alpha new end\n\nForeign pending',
    ]);
    const after = fingerprints(pkg(editor)),
      ids = groups(editor).map((g) => g.id);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(before);
    expect(await views(editor)).toEqual([
      'Alpha  end\n\nForeign',
      'Alpha old end\n\nForeign pending',
    ]);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(after);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const fresh = propertyEditor(reopened.package);
    try {
      expect(await views(fresh)).toEqual(await views(editor));
      expect(groups(fresh).map((g) => g.id)).toEqual(ids);
    } finally {
      fresh.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('imported owned addresses remain reserved before replacement allocation and across save/reopen', async () => {
  const editor = propertyEditor(fixture(false));
  try {
    const result = replaceOwnedProposals(editor, owned(editor), plan(editor), context);
    expect(result.success).toBe(true);
    if (!result.success) throw new Error(result.error);
    expect(result.ownership[0]!.native.flatMap((n) => n.addresses).map((a) => a.id)).not.toContain(
      '0'
    );
    const store = readReviewGroupStore(pkg(editor));
    expect(store.status).toBe('ready');
    if (store.status !== 'ready') throw new Error('Missing review store.');
    expect(
      [...store.retainedRecords.values()].some(
        (r) =>
          r.decoded.status === 'supported' &&
          r.decoded.value.members.some((m) => m.revision.id === '0')
      )
    ).toBe(true);
    const reopened = readOoxmlPackage(writeOoxmlPackage(pkg(editor)));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(readReviewGroupStore(reopened.package)).toMatchObject({ status: 'ready' });
  } finally {
    editor.destroy();
  }
});

test('a thrown final stage preserves exact content, metadata, references, revision, subscribers, and prior redo', () => {
  const editor = propertyEditor(fixture());
  try {
    const surface = editor.surface!,
      session = surface.session;
    const id = session.paragraphIds()[0]!;
    expect(
      surface.applyAutomationOps(() => [
        {
          op: 'insertText',
          paragraphId: id,
          offset: paragraphTextOf(session.part(), id)!.length,
          text: ' redo',
          bias: 'right',
        },
      ]).committed
    ).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const before = pkg(editor),
      fingerprint = fingerprints(before),
      revision = session.packageRevision(),
      priorGroups = groups(editor);
    let publications = 0;
    const unsubscribe = session.subscribe(() => publications++),
      apply = session.applyTreeOps.bind(session);
    const injected = spyOn(session, 'applyTreeOps').mockImplementation(
      (ops, a, h, scope, options) =>
        apply(ops, a, h, scope, {
          ...options,
          stage: (ctx) => {
            options?.stage?.(ctx);
            throw new Error('Injected replacement failure.');
          },
        })
    );
    try {
      expect(replaceOwnedProposals(editor, owned(editor), plan(editor), context)).toMatchObject({
        success: false,
        code: 'apply-failed',
      });
    } finally {
      injected.mockRestore();
      unsubscribe();
    }
    expect(pkg(editor)).toBe(before);
    expect(session.packageRevision()).toBe(revision);
    expect(fingerprints(pkg(editor))).toEqual(fingerprint);
    expect(groups(editor)).toEqual(priorGroups);
    expect(publications).toBe(0);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('Alpha old end redo');
  } finally {
    editor.destroy();
  }
});

test('owned property inverse receipts and replacement text share one history unit', async () => {
  const editor = propertyEditor(groupFixture());
  try {
    expect(executeProperties(editor).success).toBe(true);
    const selection = groups(editor)
      .filter((g) => g.properties.length)
      .map(grant);
    expect(selection).toHaveLength(1);
    const session = editor.surface!.session,
      id = groupParagraph(editor, 'The cat sat.');
    const replacement = planParagraphText(
      {
        revision: groupToken(editor),
        partName: session.part().name,
        paragraphId: id,
        text: 'The cat sat.',
      },
      'The car sat.'
    );
    if (!replacement.success) throw new Error(replacement.error);
    const before = fingerprints(pkg(editor)),
      priorViews = await views(editor);
    expect(replaceOwnedProposals(editor, selection, replacement.plan, context)).toMatchObject({
      success: true,
      committed: true,
    });
    expect(groups(editor).every((g) => g.properties.length === 0)).toBe(true);
    expect(groups(editor)).toHaveLength(1);
    expect((await views(editor))[1]).toContain('**The car sat\\.**');
    expect((await views(editor))[1]).toContain('Old activity');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(before);
    expect(await views(editor)).toEqual(priorViews);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(groups(editor).every((g) => g.properties.length === 0)).toBe(true);
  } finally {
    editor.destroy();
  }
});

for (const kind of [
  'stale',
  'incomplete',
  'wrong-before',
  'wrong-context',
  'view',
  'no-review',
] as const)
  test(`${kind} replacement refuses without changing package or history`, () => {
    const editor = propertyEditor(fixture(), kind !== 'no-review');
    try {
      const selection = owned(editor),
        candidate = plan(editor);
      if (kind === 'view') editor.surface!.setEditingMode('view');
      const before = pkg(editor),
        fingerprint = fingerprints(before),
        revision = editor.surface!.session.packageRevision();
      const result = replaceOwnedProposals(
        editor,
        kind === 'incomplete' ? [{ ...selection[0]!, native: [] }] : selection,
        kind === 'stale'
          ? { ...candidate, revision: { ...candidate.revision, revision: revision + 1 } }
          : kind === 'wrong-before'
            ? {
                ...candidate,
                operations: [
                  {
                    ...candidate.operations[0]!,
                    kind: 'text',
                    partName: editor.surface!.session.part().name,
                    paragraphId: editor.surface!.session.paragraphIds()[0]!,
                    start: 6,
                    end: 9,
                    before: 'old',
                    after: 'new',
                  },
                ],
              }
            : candidate,
        kind === 'wrong-context' ? { ...context, epoch: 'other-document' } : context
      );
      expect(result.success).toBe(false);
      expect(pkg(editor)).toBe(before);
      expect(editor.surface!.session.packageRevision()).toBe(revision);
      expect(fingerprints(pkg(editor))).toEqual(fingerprint);
    } finally {
      editor.destroy();
    }
  });

test('replacement refuses an overlapping unselected same-author proposal without changing either group', () => {
  const editor = propertyEditor(fixture());
  try {
    const session = editor.surface!.session;
    const candidate = createEditPlan(groupToken(editor), [
      {
        kind: 'text',
        partName: session.part().name,
        paragraphId: session.paragraphIds()[1]!,
        start: 9,
        end: 15,
        before: 'ending',
        after: 'ublished',
      },
    ]);
    if (!candidate.success) throw new Error(candidate.error);
    const before = pkg(editor),
      revision = session.packageRevision(),
      priorGroups = groups(editor);
    let publications = 0;
    const unsubscribe = session.subscribe(() => publications++);
    try {
      expect(replaceOwnedProposals(editor, owned(editor), candidate.plan, context)).toMatchObject({
        success: false,
        code: 'ambiguous-target',
        error: 'Replacement would change an unselected revision.',
      });
      expect(pkg(editor)).toBe(before);
      expect(session.packageRevision()).toBe(revision);
      expect(groups(editor)).toEqual(priorGroups);
      expect(publications).toBe(0);
    } finally {
      unsubscribe();
    }
  } finally {
    editor.destroy();
  }
});

test('an empty selection and replacement preserve package identity, publication count, and prior redo', () => {
  const editor = propertyEditor(fixture());
  try {
    const surface = editor.surface!,
      session = surface.session,
      paragraphId = session.paragraphIds()[0]!;
    expect(
      surface.applyAutomationOps(() => [
        { op: 'insertText', paragraphId, offset: 13, text: ' redo', bias: 'right' },
      ]).committed
    ).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const candidate = createEditPlan(groupToken(editor), []);
    if (!candidate.success) throw new Error(candidate.error);
    const before = pkg(editor),
      revision = session.packageRevision();
    let publications = 0;
    const unsubscribe = session.subscribe(() => publications++);
    try {
      expect(replaceOwnedProposals(editor, [], candidate.plan, context)).toMatchObject({
        success: true,
        committed: false,
        ownership: [],
      });
      expect(pkg(editor)).toBe(before);
      expect(session.packageRevision()).toBe(revision);
      expect(publications).toBe(0);
    } finally {
      unsubscribe();
    }
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), paragraphId)).toBe('Alpha old end redo');
  } finally {
    editor.destroy();
  }
});

test('replacement rechecks the revision after entering the native review gate', () => {
  const editor = propertyEditor(fixture());
  try {
    const surface = editor.surface!,
      session = surface.session;
    const selection = owned(editor),
      candidate = plan(editor),
      original = surface.commitReviewOps.bind(surface);
    let afterPeer: OoxmlPackage | undefined;
    let peerRevision = -1;
    const gate = spyOn(surface, 'commitReviewOps').mockImplementation((callback, kind) => {
      expect(
        surface.applyAutomationOps(() => [
          {
            op: 'insertText',
            paragraphId: session.paragraphIds()[0]!,
            offset: 13,
            text: ' peer',
            bias: 'right',
          },
        ]).committed
      ).toBe(true);
      afterPeer = pkg(editor);
      peerRevision = session.packageRevision();
      return original(callback, kind);
    });
    try {
      expect(replaceOwnedProposals(editor, selection, candidate, context)).toMatchObject({
        success: false,
        code: 'stale-revision',
      });
      if (!afterPeer) throw new Error('The native review gate did not run.');
      expect(pkg(editor)).toBe(afterPeer);
      expect(session.packageRevision()).toBe(peerRevision);
      expect(groups(editor).some((g) => g.native.some((m) => m.address.id === '0'))).toBe(true);
    } finally {
      gate.mockRestore();
    }
  } finally {
    editor.destroy();
  }
});

test('a nested replacement refuses while the outer replacement commits one history unit', () => {
  const editor = propertyEditor(fixture());
  try {
    const surface = editor.surface!,
      selection = owned(editor),
      candidate = plan(editor);
    const before = fingerprints(pkg(editor)),
      original = surface.commitReviewOps.bind(surface);
    let calls = 0;
    const gate = spyOn(surface, 'commitReviewOps').mockImplementation((callback, kind) => {
      calls++;
      expect(replaceOwnedProposals(editor, selection, candidate, context)).toMatchObject({
        success: false,
        code: 'unsupported-operation',
        error: 'A review group decision is already running.',
      });
      return original(callback, kind);
    });
    try {
      expect(replaceOwnedProposals(editor, selection, candidate, context)).toMatchObject({
        success: true,
        committed: true,
      });
      expect(calls).toBe(1);
    } finally {
      gate.mockRestore();
    }
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(surface.session.part(), surface.session.paragraphIds()[0]!)).toBe(
      'Alpha new end'
    );
  } finally {
    editor.destroy();
  }
});

test('hostile context accessors are refused without evaluation or native gate entry', () => {
  const editor = propertyEditor(fixture());
  try {
    let reads = 0;
    const hostile = {
      ...context,
      get author(): string {
        reads++;
        throw new Error('Accessor must stay inert.');
      },
    };
    const before = pkg(editor),
      revision = editor.surface!.session.packageRevision();
    const gate = spyOn(editor.surface!, 'commitReviewOps');
    try {
      expect(replaceOwnedProposals(editor, owned(editor), plan(editor), hostile)).toMatchObject({
        success: false,
        code: 'invalid-input',
      });
      expect(reads).toBe(0);
      expect(gate).not.toHaveBeenCalled();
      expect(pkg(editor)).toBe(before);
      expect(editor.surface!.session.packageRevision()).toBe(revision);
    } finally {
      gate.mockRestore();
    }
  } finally {
    editor.destroy();
  }
});

test('replacement preserves a locked content control and its owned revision', () => {
  const read = readOoxmlPackage(
    docxFixture(
      '<w:sdt><w:sdtPr><w:id w:val="42"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>' +
        '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' +
        '<w:ins w:id="0" w:author="ScioBot"><w:r><w:t>old</w:t></w:r></w:ins>' +
        '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>' +
        '</w:sdtContent></w:sdt>'
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const session = editor.surface!.session,
      before = pkg(editor),
      revision = session.packageRevision();
    const selection = owned(editor),
      candidate = plan(editor),
      priorGroups = groups(editor);
    let publications = 0;
    const unsubscribe = session.subscribe(() => publications++);
    try {
      expect(replaceOwnedProposals(editor, selection, candidate, context).success).toBe(false);
      expect(pkg(editor)).toBe(before);
      expect(session.packageRevision()).toBe(revision);
      expect(groups(editor)).toEqual(priorGroups);
      expect(publications).toBe(0);
    } finally {
      unsubscribe();
    }
  } finally {
    editor.destroy();
  }
});
