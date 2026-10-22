import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import {
  blockAst,
  blockContext,
  blockEditor,
  blockP,
  blockToken,
  blockViews,
  decideBlocks,
  manualBlockPlan,
  targetBlockPlan,
} from '../../test/support/block-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackSiblingPlan } from './track-blocks';

for (const index of [2, 5])
  for (const decision of ['acceptRevision', 'rejectRevision'] as const)
    test(`C01 ${pairs.cases[index]!.name}: tracked ${decision} preserves native siblings, history, and saved receipts`, async () => {
      const fixture = pairs.cases[index]!,
        mounted = blockEditor(fixture.beforeMarkdown.split('\n\n').map(blockP).join('')),
        editor = mounted.editor;
      try {
        const part = editor.surface!.session.part(),
          ids = editor.surface!.session.paragraphIds();
        const plan = await targetBlockPlan(editor, fixture.afterMarkdown);
        expect(plan.operations.every((edit) => edit.kind === 'siblings')).toBe(true);
        const removed = plan.operations.flatMap((edit) =>
          edit.kind === 'siblings' ? edit.removedIds : []
        );
        const kept = ids.filter((id) => !removed.includes(id));
        let publications = 0;
        const stop = editor.surface!.session.subscribe(() => publications++),
          token = blockToken(editor);
        const result = trackSiblingPlan(editor, plan, blockContext);
        stop();
        if (!result.success) throw new Error(result.error);
        expect(publications).toBe(1);
        expect(result.revision.revision).toBe(token.revision + 1);
        expect(result.proposals).toHaveLength(1);
        expect(result.proposals[0]!.addresses).toHaveLength(1);
        for (const id of kept)
          expect(findNode(editor.surface!.session.part(), id)).toBe(findNode(part, id));
        for (const site of collectRevisionSites(editor.surface!.session.part())) {
          expect(
            site.node.attributes.find((attribute) => attribute.localName === 'author')?.value
          ).toBe(blockContext.author);
          expect(
            site.node.attributes.find((attribute) => attribute.localName === 'date')?.value
          ).toBe('2026-10-19T09:00:00.000Z');
        }
        expect((await blockViews(editor)).map(blockAst)).toEqual([
          blockAst(fixture.beforeMarkdown),
          blockAst(fixture.afterMarkdown),
        ]);
        const pending = new Uint8Array(await editor.save());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect((await blockViews(editor)).map(blockAst)).toEqual([
          blockAst(fixture.beforeMarkdown),
          blockAst(fixture.beforeMarkdown),
        ]);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect((await blockViews(editor)).map(blockAst)).toEqual([
          blockAst(fixture.beforeMarkdown),
          blockAst(fixture.afterMarkdown),
        ]);
        decideBlocks(editor, result.proposals, decision);
        const expected =
          decision === 'acceptRevision' ? fixture.afterMarkdown : fixture.beforeMarkdown;
        expect((await blockViews(editor)).map(blockAst)).toEqual([
          blockAst(expected),
          blockAst(expected),
        ]);
        for (const id of kept)
          expect(findNode(editor.surface!.session.part(), id)).toBe(findNode(part, id));
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect((await blockViews(editor)).map(blockAst)).toEqual([
          blockAst(fixture.beforeMarkdown),
          blockAst(fixture.afterMarkdown),
        ]);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect((await blockViews(editor)).map(blockAst)).toEqual([
          blockAst(expected),
          blockAst(expected),
        ]);
        const container = document.createElement('div'),
          reopened = createDocxEditor({
            container,
            document: pending,
            modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
          });
        try {
          expect((await blockViews(reopened)).map(blockAst)).toEqual([
            blockAst(fixture.beforeMarkdown),
            blockAst(fixture.afterMarkdown),
          ]);
          decideBlocks(reopened, result.proposals, decision);
          expect((await blockViews(reopened)).map(blockAst)).toEqual([
            blockAst(expected),
            blockAst(expected),
          ]);
        } finally {
          reopened.destroy();
        }
      } finally {
        mounted.destroy();
      }
    }, 30_000);

for (const at of [0, 1, 3])
  test(`new paragraph insertion at ${at} tracks only new content`, async () => {
    const mounted = blockEditor(['One', 'Two', 'Three'].map(blockP).join('')),
      editor = mounted.editor;
    try {
      const ids = editor.surface!.session.paragraphIds(),
        before = editor.surface!.session.part();
      const result = trackSiblingPlan(
        editor,
        manualBlockPlan(editor, at, 0, 'Added'),
        blockContext
      );
      if (!result.success) throw new Error(result.error);
      for (const id of ids)
        expect(findNode(editor.surface!.session.part(), id)).toBe(findNode(before, id));
      const expected = ['One', 'Two', 'Three'];
      expected.splice(at, 0, 'Added');
      expect(await blockViews(editor)).toEqual(['One\n\nTwo\n\nThree', expected.join('\n\n')]);
      decideBlocks(editor, result.proposals, 'rejectRevision');
      expect(editor.surface!.session.paragraphIds()).toEqual(ids);
      expect(await blockViews(editor)).toEqual(['One\n\nTwo\n\nThree', 'One\n\nTwo\n\nThree']);
    } finally {
      mounted.destroy();
    }
  });

for (const at of [0, 1, 2])
  test(`whole paragraph deletion at ${at} preserves untouched neighbor IDs`, async () => {
    const mounted = blockEditor(['One', 'Two', 'Three'].map(blockP).join('')),
      editor = mounted.editor;
    try {
      const ids = editor.surface!.session.paragraphIds(),
        before = editor.surface!.session.part();
      const result = trackSiblingPlan(editor, manualBlockPlan(editor, at, 1), blockContext);
      if (!result.success) throw new Error(result.error);
      const expected = ['One', 'Two', 'Three'];
      expected.splice(at, 1);
      expect(await blockViews(editor)).toEqual(['One\n\nTwo\n\nThree', expected.join('\n\n')]);
      decideBlocks(editor, result.proposals, 'acceptRevision');
      expect(editor.surface!.session.paragraphIds()).toEqual(
        ids.filter((_, index) => index !== at)
      );
      for (const id of ids.filter((_, index) => index !== at))
        expect(findNode(editor.surface!.session.part(), id)).toBe(findNode(before, id));
    } finally {
      mounted.destroy();
    }
  });

test('empty plans leave document revision and history unchanged', () => {
  const mounted = blockEditor(blockP('Kept'));
  try {
    const editor = mounted.editor,
      token = blockToken(editor),
      pkg = editor.surface!.session.currentPackage();
    const result = trackSiblingPlan(
      editor,
      { schemaVersion: 1, revision: token, operations: [] },
      blockContext
    );
    expect(result).toMatchObject({ success: true, revision: token, proposals: [] });
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
      canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)
    );
    expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(
      writeOoxmlPackage(pkg)
    );
    expect(editor.surface!.state().canUndo).toBe(false);
    expect(
      paragraphTextOf(editor.surface!.session.part(), editor.surface!.session.paragraphIds()[0]!)
    ).toBe('Kept');
  } finally {
    mounted.destroy();
  }
});

test('one gap can replace several paragraphs with distinct native insertion/deletion addresses', async () => {
  const mounted = blockEditor(['Kept', 'Old first', 'Old second', 'Tail'].map(blockP).join(''));
  try {
    const editor = mounted.editor;
    const result = trackSiblingPlan(
      editor,
      manualBlockPlan(editor, 1, 2, '## New\n\nNew body'),
      blockContext
    );
    if (!result.success) throw new Error(result.error);
    expect(result.proposals[0]!.addresses).toHaveLength(2);
    expect(await blockViews(editor)).toEqual([
      'Kept\n\nOld first\n\nOld second\n\nTail',
      'Kept\n\n## New\n\nNew body\n\nTail',
    ]);
    decideBlocks(editor, result.proposals, 'acceptRevision');
    expect(await blockViews(editor)).toEqual([
      'Kept\n\n## New\n\nNew body\n\nTail',
      'Kept\n\n## New\n\nNew body\n\nTail',
    ]);
  } finally {
    mounted.destroy();
  }
});

test('new paragraph marks and native links belong to its insertion, including saved decisions', async () => {
  const mounted = blockEditor(blockP('Kept'));
  try {
    const editor = mounted.editor,
      target =
        'New **bold** *italic* ~~strike~~ [link](https://example.test/path?x=1&y=2 "A title") and [anchor](#here)';
    const result = trackSiblingPlan(editor, manualBlockPlan(editor, 1, 0, target), blockContext);
    if (!result.success) throw new Error(result.error);
    expect((await blockViews(editor)).map(blockAst)).toEqual([
      blockAst('Kept'),
      blockAst(`Kept\n\n${target}`),
    ]);
    expect(result.proposals[0]!.addresses).toHaveLength(1);
    const bytes = new Uint8Array(await editor.save()),
      reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
    try {
      decideBlocks(reopened, result.proposals, 'acceptRevision');
      expect((await blockViews(reopened)).map(blockAst)).toEqual([
        blockAst(`Kept\n\n${target}`),
        blockAst(`Kept\n\n${target}`),
      ]);
    } finally {
      reopened.destroy();
    }
    decideBlocks(editor, result.proposals, 'rejectRevision');
    expect(await blockViews(editor)).toEqual(['Kept', 'Kept']);
  } finally {
    mounted.destroy();
  }
});

test('one request removes the sole paragraph through the native required-paragraph fallback', async () => {
  const mounted = blockEditor(blockP('Only'));
  try {
    const editor = mounted.editor,
      result = trackSiblingPlan(editor, manualBlockPlan(editor, 0, 1), blockContext);
    if (!result.success) throw new Error(result.error);
    expect(await blockViews(editor)).toEqual(['Only', '']);
    decideBlocks(editor, result.proposals, 'acceptRevision');
    expect(await blockViews(editor)).toEqual(['', '']);
    expect(editor.surface!.session.paragraphIds()).toHaveLength(1);
  } finally {
    mounted.destroy();
  }
});

for (const scenario of [
  'stale',
  'epoch',
  'document',
  'view',
  'module',
  'anchor',
  'window',
  'duplicate',
  'mixed',
  'unsafe-link',
  'metadata',
  'pending',
  'section',
] as const)
  test(`invalid or unsupported sibling input refuses without mutation: ${scenario}`, () => {
    const mounted = blockEditor(
      ['One', 'Two', 'Three'].map(blockP).join(''),
      scenario !== 'module'
    );
    try {
      const editor = mounted.editor,
        surface = editor.surface!,
        ids = surface.session.paragraphIds();
      let plan = manualBlockPlan(editor, 1, 1, 'New');
      const edit = plan.operations[0]!;
      if (edit.kind !== 'siblings') throw new Error('Missing sibling gap');
      if (scenario === 'stale')
        plan = { ...plan, revision: { ...plan.revision, revision: plan.revision.revision + 1 } };
      if (scenario === 'epoch' || scenario === 'document')
        plan = {
          ...plan,
          revision: {
            ...plan.revision,
            [scenario === 'epoch' ? 'epoch' : 'documentId']: 'foreign',
          },
        };
      if (scenario === 'view') editor.setEditingMode('viewing');
      if (scenario === 'anchor') plan = { ...plan, operations: [{ ...edit, beforeId: 'missing' }] };
      if (scenario === 'window') plan = { ...plan, operations: [{ ...edit, removedIds: [] }] };
      if (scenario === 'duplicate') plan = { ...plan, operations: [edit, edit] };
      if (scenario === 'mixed')
        plan = {
          ...plan,
          operations: [
            ...plan.operations,
            {
              kind: 'text',
              partName: edit.partName,
              paragraphId: ids[0]!,
              start: 0,
              end: 0,
              before: '',
              after: 'x',
            },
          ],
        };
      if (scenario === 'unsafe-link')
        plan = {
          ...plan,
          operations: [
            {
              ...edit,
              inserted: [
                {
                  type: 'paragraph',
                  children: [
                    {
                      type: 'link',
                      url: 'javascript:alert(1)',
                      children: [{ type: 'text', value: 'bad' }],
                    },
                  ],
                },
              ],
            },
          ],
        };
      if (scenario === 'metadata')
        plan = {
          ...plan,
          operations: [
            {
              ...edit,
              inserted: [
                {
                  type: 'paragraph',
                  data: { unknown: 'keep' },
                  children: [{ type: 'text', value: 'New' }],
                },
              ],
            },
          ],
        };
      if (scenario === 'pending') {
        surface.applyAutomationOps(
          () => [
            {
              op: 'insertText',
              paragraphId: ids[1]!,
              offset: 0,
              text: 'pending',
              revision: { author: 'Other' },
            },
          ],
          { kind: 'body' },
          undefined,
          true,
          true
        );
        plan = { ...plan, revision: blockToken(editor) };
      }
      if (scenario === 'section') {
        surface.applyAutomationOps(
          () => [{ op: 'setSectionMark', paragraphId: ids[1]! }],
          { kind: 'body' },
          undefined,
          false,
          true
        );
        plan = { ...plan, revision: blockToken(editor) };
      }
      const before = writeOoxmlPackage(surface.session.currentPackage()),
        token = blockToken(editor),
        canUndo = surface.state().canUndo;
      expect(trackSiblingPlan(editor, plan, blockContext).success).toBe(false);
      expect(writeOoxmlPackage(surface.session.currentPackage())).toEqual(before);
      expect(blockToken(editor)).toEqual(token);
      expect(surface.state().canUndo).toBe(canUndo);
    } finally {
      mounted.destroy();
    }
  });

test('a failed post-operation hook rolls back both structural revisions and allocated links', () => {
  const mounted = blockEditor(blockP('Kept'));
  try {
    const editor = mounted.editor,
      surface = editor.surface!,
      original = surface.applyAutomationOps.bind(surface);
    const plan = manualBlockPlan(editor, 1, 0, '[new](https://example.test/new)'),
      before = writeOoxmlPackage(surface.session.currentPackage()),
      token = blockToken(editor);
    surface.applyAutomationOps = (staged, scope, edits, review, separate, phase) =>
      original(
        staged,
        scope,
        edits?.map((hook) => (pkg) => {
          hook(pkg);
          throw new Error('injected post-operation failure');
        }),
        review,
        separate,
        phase
      );
    try {
      expect(trackSiblingPlan(editor, plan, blockContext)).toMatchObject({
        success: false,
        code: 'apply-failed',
      });
      expect(writeOoxmlPackage(surface.session.currentPackage())).toEqual(before);
      expect(blockToken(editor)).toEqual(token);
      expect(surface.state().canUndo).toBe(false);
    } finally {
      surface.applyAutomationOps = original;
    }
  } finally {
    mounted.destroy();
  }
});

test('a same-author neighboring proposal keeps its own native addresses and pending content', async () => {
  const mounted = blockEditor(['Prior', 'Remove', 'Tail'].map(blockP).join(''));
  try {
    const editor = mounted.editor,
      surface = editor.surface!,
      ids = surface.session.paragraphIds();
    const revision = { author: blockContext.author, date: '2026-10-19T09:00:00.000Z' };
    surface.applyAutomationOps(
      () => [
        { op: 'deleteText', paragraphId: ids[0]!, start: 0, end: 5, revision },
        { op: 'setParagraphMarkRevision', paragraphId: ids[0]!, kind: 'del', revision },
      ],
      { kind: 'body' },
      undefined,
      true,
      true
    );
    const prior = collectRevisionSites(surface.session.part()).map((site) => [
      site.node.id,
      JSON.stringify(site.node),
    ]);
    const result = trackSiblingPlan(editor, manualBlockPlan(editor, 1, 1), blockContext);
    if (!result.success) throw new Error(result.error);
    for (const [id, value] of prior)
      expect(JSON.stringify(findNode(surface.session.part(), id!))).toBe(value);
    decideBlocks(editor, result.proposals, 'acceptRevision');
    for (const [id, value] of prior)
      expect(JSON.stringify(findNode(surface.session.part(), id!))).toBe(value);
    expect(await blockViews(editor)).toEqual(['Prior\n\nTail', 'Tail']);
  } finally {
    mounted.destroy();
  }
});
