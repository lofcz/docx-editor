import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { groupToken, groups } from '../../test/support/review-groups';
import { targetAstKey } from '../diff/target-key';
import { parseTargetMarkdown } from '../projection/parse-target';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals } from './recompose';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
type Wave = Extract<Awaited<ReturnType<typeof recomposeOwnedProposals>>, { success: true }>;
const context = {
  documentId: 'property-records',
  epoch: 'load-1',
  author: 'Bot',
  date: '2026-10-21T14:00:00.000Z',
};
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>' +
  `<w:tc><w:tcPr/>${p('Cell one')}</w:tc><w:tc><w:tcPr/>${p('Cell two')}</w:tc></w:tr></w:tbl>`;
function options(withTable = false) {
  return {
    document: docxFixture(p('Alpha end') + p('Keep') + (withTable ? table : '') + p('Tail')),
    documentId: context.documentId,
    ownedAttributes: true,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  };
}
function ast(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
}
async function views(editor: DocxEditorInstance) {
  const [before, after] = await Promise.all([
    readCommittedView(editor.surface!.session, groupToken(editor)),
    readWorkingView(editor.surface!.session, groupToken(editor)),
  ]);
  if (!before.success || !after.success) throw new Error('Missing mixed shared views.');
  return [before.view.projection.text, after.view.projection.text];
}
async function wave(editor: DocxEditorInstance, targetMarkdown: string, previous?: Wave) {
  const revision = groupToken(editor);
  const publications: { origin: string; fingerprint: string }[] = [];
  const unsubscribe = editor.surface!.session.subscribe((change) => {
    const source = editor.surface!.session.currentPackage();
    publications.push({
      origin: change.origin,
      fingerprint: JSON.stringify(
        [...source.parts]
          .map(([name, part]) => {
            // OPC relationships form a keyed set. Room reflection can change their XML order.
            const canonical =
              name.endsWith('.rels') && part.root.kind === 'generic'
                ? {
                    ...part,
                    root: {
                      ...part.root,
                      children: [...part.root.children].sort((a, b) =>
                        canonicalOoxmlFingerprint(a).localeCompare(canonicalOoxmlFingerprint(b))
                      ),
                    },
                  }
                : part;
            return [name, canonicalOoxmlFingerprint(canonical)];
          })
          .sort(([a], [b]) => a!.localeCompare(b!))
      ),
    });
  });
  try {
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision,
        targetMarkdown,
        ownership: previous?.ownership ?? [],
        contentProofs: previous?.contentProofs ?? [],
      },
      { ...context, epoch: revision.epoch }
    );
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    expect(result.committed).toBe(true);
    expect(publications.filter((entry) => !entry.origin.endsWith('.remote'))).toHaveLength(1);
    expect(new Set(publications.map((entry) => entry.fingerprint)).size).toBe(1);
    return result;
  } finally {
    unsubscribe();
  }
}
function append(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const offset = paragraphTextOf(editor.surface!.session.part(), paragraphId)!.length;
  expect(
    editor.surface!.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId, offset, text, bias: 'right' }],
      { kind: 'body' },
      [],
      false,
      true
    )
  ).toMatchObject({ committed: true, rejected: false });
}
function unnested(editor: DocxEditorInstance) {
  const visit = (node: OoxmlNode, inside = false) => {
    if (node.kind === 'textValue') return;
    const contentRevision = node.kind === 'revisionInsert' || node.kind === 'revisionDelete';
    if (contentRevision) expect(inside).toBe(false);
    for (const child of node.children) visit(child, inside || contentRevision);
  };
  visit(editor.surface!.session.part().root);
  const sites = collectRevisionSites(editor.surface!.session.part());
  expect(new Set(sites.map((site) => site.node.id)).size).toBe(sites.length);
}
async function check(pair: Pair, original: string, proposed: string) {
  for (const editor of pair.editors) {
    expect((await views(editor)).map(ast)).toEqual([original, proposed].map(ast));
    unnested(editor);
  }
}
async function geometry(pair: Pair, id: string) {
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
  const offset = paragraphTextOf(bob.surface!.session.part(), id)!.length;
  bob.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find((entry) => entry.actorId === 'bob');
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const a = caretAt(alice.surface!.layout(), { paragraphId: id, offset });
  const b = caretAt(bob.surface!.layout(), { paragraphId: id, offset });
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(pair: Pair) {
  const expected = await views(pair.editors[0]!);
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  const expectedGroups = groups(pair.editors[0]!).map((group) => group.id);
  for (const room of [false, true]) {
    const reopened = await createTwoEditorFixture({
      ...options(),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: room ? [401, 433] : [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of reopened.editors) {
        expect((await views(editor)).map(ast)).toEqual(expected.map(ast));
        expect(groups(editor).map((group) => group.id)).toEqual(expectedGroups);
      }
    } finally {
      reopened.destroy();
    }
  }
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const) {
    test(`mixed waves preserve independent peer text, both histories, geometry, and reopening; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const [title, keep, tail] = alice.surface!.session.paragraphIds();
        const firstTarget =
          '**Alpha XYZ end**\n\nKeep\n\nDraft **bold** and `code` [link](https://example.com/one)\n\nTail';
        const first = await wave(alice, firstTarget);
        await pair.settle();
        pair.pause();
        const target =
          '*Alpha Q end*\n\nKeep\n\nFinal *italic* [link](https://example.com/two)\n\nTail';
        await wave(alice, target, first);
        append(bob, keep!, ' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        const original = 'Alpha end\n\nKeep Peer\n\nTail';
        const working = target.replace('Keep', 'Keep Peer');
        await check(pair, original, working);
        for (const editor of pair.editors) {
          expect(findNode(editor.surface!.session.part(), title!)).not.toBeNull();
          expect(findNode(editor.surface!.session.part(), keep!)).not.toBeNull();
          expect(findNode(editor.surface!.session.part(), tail!)).not.toBeNull();
        }
        for (let cycle = 0; cycle < 2; cycle++) {
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, original, firstTarget.replace('Keep', 'Keep Peer'));
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, original, working);
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, 'Alpha end\n\nKeep\n\nTail', target);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original, working);
        const inserted = alice
          .surface!.session.paragraphIds()
          .find((id) =>
            paragraphTextOf(alice.surface!.session.part(), id)?.includes('Final italic')
          );
        expect(inserted).toBeDefined();
        await geometry(pair, inserted!);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);

    test(`mixed cell and body marks retain offline neighboring cells and complete undo; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(true), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const cell = alice
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Cell two')!;
        const firstTarget =
          '**Alpha XYZ end**\n\nKeep\n\n| Cell **ONE** | Cell two |\n| --- | --- |\n\nTail';
        const first = await wave(alice, firstTarget);
        await pair.settle();
        const originalBefore = (await views(alice))[0]!;
        pair.pause();
        const target = '*Alpha Q end*\n\nKeep\n\n| Cell *new* | Cell two |\n| --- | --- |\n\nTail';
        await wave(alice, target, first);
        append(bob, cell, ' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        const original = originalBefore.replace('Cell two', 'Cell two Peer');
        await check(pair, original, target.replace('Cell two', 'Cell two Peer'));
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original, firstTarget.replace('Cell two', 'Cell two Peer'));
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original, target.replace('Cell two', 'Cell two Peer'));
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, originalBefore, target);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await geometry(pair, cell);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);

    test(`deleting a mixed wave owner offline preserves other inserted content; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const title = alice.surface!.session.paragraphIds()[0]!;
        const first = await wave(alice, '**Alpha XYZ end**\n\nKeep\n\nDraft\n\nTail');
        await pair.settle();
        pair.pause();
        await wave(alice, '*Alpha Q end*\n\nKeep\n\nFinal\n\nTail', first);
        expect(
          bob.surface!.applyAutomationOps(
            () => [{ op: 'deleteBlock', blockId: title }],
            { kind: 'body' },
            [],
            false,
            true
          ).committed
        ).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        await check(pair, 'Keep\n\nTail', 'Keep\n\nFinal\n\nTail');
        for (const editor of pair.editors)
          expect(findNode(editor.surface!.session.part(), title)).toBeNull();
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, 'Alpha end\n\nKeep\n\nTail', '*Alpha Q end*\n\nKeep\n\nFinal\n\nTail');
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, 'Keep\n\nTail', 'Keep\n\nFinal\n\nTail');
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);
    test(`mixed decisions and manual suffixes survive the next complete target; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const [, keep, tail] = alice.surface!.session.paragraphIds();
        const first = await wave(alice, '**Alpha XYZ end**\n\n*Keep*\n\nDraft\n\nTail');
        await pair.settle();
        const draft = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Draft');
        expect(draft).toBeDefined();
        for (const [id, decide] of [
          [keep, acceptReviewGroup],
          [draft, rejectReviewGroup],
        ] as const) {
          const selected = groups(bob).filter((group) => group.anchor?.paragraphId === id);
          expect(selected.length).toBeGreaterThan(0);
          for (const group of selected) {
            const token = groupToken(bob);
            const result = decide(bob, group.id, token, token);
            if (!result.success) throw new Error(result.error);
            await pair.settle();
          }
        }
        pair.pause();
        append(bob, tail!, ' Manual');
        pair.reconnect(receiver);
        await pair.settle();
        const target = '*Alpha Q end*\n\n*Keep*\n\nFinal\n\nTail Manual';
        await wave(alice, target, first);
        await pair.settle();
        const original = 'Alpha end\n\n*Keep*\n\nTail Manual';
        await check(pair, original, target);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original, '**Alpha XYZ end**\n\n*Keep*\n\nTail Manual');
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original, target);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);

    test(`observed peer property edits invalidate owned format evidence without writes; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const title = alice.surface!.session.paragraphIds()[0]!;
        const first = await wave(alice, '**Alpha end**\n\nKeep\n\nTail');
        await pair.settle();
        pair.pause();
        expect(
          bob.surface!.applyAutomationOps(
            () => [
              {
                op: 'setRunProperties',
                paragraphId: title,
                start: 0,
                end: 9,
                properties: [{ localName: 'i' }],
              },
            ],
            { kind: 'body' },
            [],
            false,
            true
          ).committed
        ).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        const session = alice.surface!.session,
          before = session.currentPackage(),
          revision = groupToken(alice);
        const prior = groups(alice);
        let publications = 0;
        const stop = session.subscribe(() => publications++);
        try {
          expect(
            await recomposeOwnedProposals(
              alice,
              {
                revision,
                targetMarkdown: '*Alpha Q end*\n\nKeep\n\nTail',
                ownership: first.ownership,
                contentProofs: first.contentProofs,
              },
              { ...context, epoch: revision.epoch }
            )
          ).toMatchObject({ success: false, code: 'ambiguous-target' });
          expect(session.currentPackage()).toBe(before);
          expect(groupToken(alice)).toEqual(revision);
          expect(groups(alice)).toEqual(prior);
          expect(publications).toBe(0);
        } finally {
          stop();
        }
        await check(pair, 'Alpha end\n\nKeep\n\nTail', '*Alpha end*\n\nKeep\n\nTail');
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, 'Alpha end\n\nKeep\n\nTail', '**Alpha end**\n\nKeep\n\nTail');
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);

    test(`peer paragraph formatting inside an owned insertion refuses a later wave; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const target = 'Alpha end\n\nKeep\n\nDraft\n\nTail';
        const first = await wave(alice, target);
        await pair.settle();
        const draft = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Draft')!;
        expect(draft).toBeDefined();
        pair.pause();
        expect(
          bob.surface!.applyAutomationOps(
            () => [
              {
                op: 'setParagraphProperties',
                paragraphId: draft,
                properties: [{ localName: 'jc', attributes: { val: 'center' } }],
              },
            ],
            { kind: 'body' },
            [],
            false,
            true
          ).committed
        ).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors)
          expect(directParagraphProperties(editor.surface!.session.part(), draft)).toContainEqual({
            localName: 'jc',
            attributes: { val: 'center' },
          });
        await geometry(pair, draft);
        const session = alice.surface!.session,
          before = session.currentPackage(),
          revision = groupToken(alice),
          prior = groups(alice);
        let publications = 0;
        const stop = session.subscribe(() => publications++);
        try {
          expect(
            await recomposeOwnedProposals(
              alice,
              {
                revision,
                targetMarkdown: target.replace('Draft', 'Final'),
                ownership: first.ownership,
                contentProofs: first.contentProofs,
              },
              { ...context, epoch: revision.epoch }
            )
          ).toMatchObject({ success: false, code: 'ambiguous-target' });
          expect(session.currentPackage()).toBe(before);
          expect(groupToken(alice)).toEqual(revision);
          expect(groups(alice)).toEqual(prior);
          expect(publications).toBe(0);
        } finally {
          stop();
        }
        await check(pair, 'Alpha end\n\nKeep\n\nTail', target);
        await saved(pair);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(
            directParagraphProperties(editor.surface!.session.part(), draft).some(
              (property) => property.localName === 'jc'
            )
          ).toBe(false);
        await wave(alice, target.replace('Draft', 'Final'), first);
        await pair.settle();
        await check(pair, 'Alpha end\n\nKeep\n\nTail', target.replace('Draft', 'Final'));
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, 'Alpha end\n\nKeep\n\nTail', target);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, 'Alpha end\n\nKeep\n\nTail', target.replace('Draft', 'Final'));
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);

    test(`mixed retained hyperlink waves preserve titles, peer text, decisions, and history; ${clientIds}/${receiver}`, async () => {
      const document = docxFixture(
        p('Alpha end') +
          p('Keep') +
          '<w:p><w:hyperlink r:id="old" w:tooltip="old title"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>' +
          p('Tail'),
        {
          relationships:
            '<Relationship Id="old" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/old" TargetMode="External"/>',
        }
      );
      const pair = await createTwoEditorFixture({ ...options(), document, clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const [, keep, linkParagraph] = alice.surface!.session.paragraphIds();
        const firstTarget =
          '**Alpha XYZ end**\n\nKeep\n\n[Link](https://example.com/new "new title")\n\nTail';
        const first = await wave(alice, firstTarget);
        await pair.settle();
        pair.pause();
        const finalTarget =
          '*Alpha Q end*\n\nKeep\n\n[Link](https://example.com/final "final title")\n\nTail';
        const second = await wave(alice, finalTarget, first);
        append(bob, keep!, ' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        const original =
          'Alpha end\n\nKeep Peer\n\n[Link](https://example.com/old "old title")\n\nTail';
        const working = finalTarget.replace('Keep', 'Keep Peer');
        await check(pair, original, working);
        await geometry(pair, linkParagraph!);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original, firstTarget.replace('Keep', 'Keep Peer'));
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original, working);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(pair, original.replace('Keep Peer', 'Keep'), finalTarget);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const selected = groups(bob).filter((group) => group.anchor?.paragraphId === linkParagraph);
        expect(selected).toHaveLength(1);
        const decision = acceptReviewGroup(bob, selected[0]!.id, groupToken(bob), groupToken(bob));
        if (!decision.success) throw new Error(decision.error);
        await pair.settle();
        const committed = original.replace(
          'https://example.com/old "old title"',
          'https://example.com/final "final title"'
        );
        await check(pair, committed, working);
        await wave(alice, working.replace('Alpha Q', 'Alpha R'), second);
        await pair.settle();
        await check(pair, committed, working.replace('Alpha Q', 'Alpha R'));
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30_000);
  }
