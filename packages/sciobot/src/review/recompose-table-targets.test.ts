import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  collectRevisionSites,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import { groupToken, groups } from '../../test/support/review-groups';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
import { recomposeOwnedProposals } from './recompose';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

type Editor = ReturnType<typeof propertyEditor>;
type Wave = Extract<Awaited<ReturnType<typeof recomposeOwnedProposals>>, { success: true }>;
const context = { ...propertyContext, date: '2026-10-21T18:00:00.000Z' };
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
const input = () =>
  docxFixture(
    p('Alpha end') +
      '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      [
        ['H1', 'H2'],
        ['A', 'B'],
        ['C', 'D'],
      ]
        .map((row) => '<w:tr>' + row.map(cell).join('') + '</w:tr>')
        .join('') +
      '</w:tbl>' +
      p('Tail')
  );
const htmlTarget =
  '**Alpha XYZ end**\n\n<table><tr><td colspan="2">\n\nTitle\n\n</td></tr><tr><td>\n\nA\n\n</td><td>\n\nB\n\n</td></tr><tr><td>\n\nC\n\n</td><td>\n\nD\n\n</td></tr></table>\n\nTail';
function ast(text: string) {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw new Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
}
function editor_() {
  const read = readOoxmlPackage(input());
  if (!read.ok) throw new Error(read.reason);
  return propertyEditor(read.package);
}
const fingerprints = (editor: Editor) =>
  [...editor.surface!.session.currentPackage().parts].map(([name, part]) => [
    name,
    canonicalOoxmlFingerprint(part),
  ]);
async function working(editor: Editor, auto = true) {
  const result = await readWorkingView(
    editor.surface!.session,
    groupToken(editor),
    auto ? { tableSyntax: 'auto' } : {}
  );
  if (!result.success) throw new Error(result.error);
  return result.view.projection.text;
}
async function committed(editor: Editor) {
  const result = await readCommittedView(editor.surface!.session, groupToken(editor), {
    tableSyntax: 'auto',
  });
  if (!result.success) throw new Error(result.error);
  return result.view.projection.text;
}
async function wave(editor: Editor, targetMarkdown: string, previous?: Wave, expected = true) {
  let count = 0;
  const stop = editor.surface!.session.subscribe(() => count++);
  try {
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision: groupToken(editor),
        targetMarkdown,
        ownership: previous?.ownership ?? [],
        contentProofs: previous?.contentProofs ?? [],
      },
      context
    );
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    expect(result.committed).toBe(expected);
    expect(count).toBe(expected ? 1 : 0);
    return result;
  } finally {
    stop();
  }
}

test('unchanged default GFM retains an owned merge, exact ledger, package reference, and prior redo', async () => {
  const editor = editor_();
  try {
    const first = await wave(editor, htmlTarget),
      session = editor.surface!.session;
    const tail = session.paragraphIds().at(-1)!;
    expect(
      editor.surface!.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' manual' }],
        { kind: 'body' },
        [],
        false,
        true,
        'after'
      ).committed
    ).toBe(true);
    const manual = fingerprints(editor);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const before = fingerprints(editor),
      pkg = session.currentPackage();
    const result = await wave(editor, await working(editor, false), first, false);
    expect(result.ownership).toEqual(first.ownership);
    expect(result.contentProofs).toEqual(first.contentProofs);
    expect(session.currentPackage()).toBe(pkg);
    expect(fingerprints(editor)).toEqual(before);
    expect(ast(await working(editor))).toBe(ast(htmlTarget));
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(fingerprints(editor)).toEqual(manual);
  } finally {
    editor.destroy();
  }
}, 30000);

for (const change of ['body', 'head', 'both', 'split'] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`default GFM ${change} after owned merge retains native history, reopening, and ${decision}`, async () => {
      const editor = editor_();
      try {
        const original = await committed(editor),
          first = await wave(editor, htmlTarget),
          before = fingerprints(editor);
        let target = await working(editor, false),
          expected = htmlTarget;
        if (change === 'body' || change === 'both') {
          target = target.replace('Alpha XYZ', 'Alpha Q');
          expected = expected.replace('Alpha XYZ', 'Alpha Q');
        }
        if (change === 'head' || change === 'both') {
          target = target.replace('Title', 'Final **bold**');
          expected = expected.replace('Title', 'Final **bold**');
        }
        if (change === 'split') {
          target = target.replace('| Title |  |', '| Title | New **cell** |');
          expected = target;
        }
        await wave(editor, target, first);
        expect(ast(await working(editor))).toBe(ast(expected));
        expect(ast(await committed(editor))).toBe(ast(original));
        const after = fingerprints(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprints(editor)).toEqual(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprints(editor)).toEqual(after);
        for (const part of editor.surface!.session.currentPackage().parts.values())
          expect(collectRevisionSites(part).every((site) => !site.refused)).toBe(true);
        const read = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!read.ok) throw new Error(read.reason);
        const reopened = propertyEditor(read.package);
        try {
          expect(ast(await working(reopened))).toBe(ast(expected));
          expect(groups(reopened).map((g) => g.id)).toEqual(groups(editor).map((g) => g.id));
        } finally {
          reopened.destroy();
        }
        for (const group of [...groups(editor)]) {
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            group.id,
            groupToken(editor),
            groupToken(editor)
          );
          if (!result.success) throw new Error(result.error);
        }
        expect(ast(await committed(editor))).toBe(ast(decision === 'accept' ? expected : original));
        expect(ast(await working(editor))).toBe(ast(decision === 'accept' ? expected : original));
      } finally {
        editor.destroy();
      }
    }, 30000);

test('working geometry normalization retains same-author foreign cell proposals and their exact native identities', async () => {
  const editor = editor_();
  try {
    const foreignTarget = (await working(editor, false)).replace('| C | D |', '| C Foreign | D |');
    const foreign = await wave(editor, foreignTarget);
    const foreignIds = new Set(
      groups(editor).flatMap((group) => group.native.flatMap((member) => member.siteNodeIds))
    );
    const before = new Map(
      collectRevisionSites(editor.surface!.session.part())
        .filter((site) => foreignIds.has(site.node.id))
        .map((site) => [site.node.id, JSON.stringify(site.node)])
    );
    expect(before.size).toBeGreaterThan(0);
    const first = await wave(editor, htmlTarget.replace('C\n\n', 'C Foreign\n\n'));
    const target = (await working(editor, false))
      .replace('Title', 'Final')
      .replace('Alpha XYZ', 'Alpha Q');
    await wave(editor, target, first);
    expect(ast(await working(editor))).toBe(
      ast(
        htmlTarget
          .replace('Title', 'Final')
          .replace('Alpha XYZ', 'Alpha Q')
          .replace('C\n\n', 'C Foreign\n\n')
      )
    );
    expect(
      new Map(
        collectRevisionSites(editor.surface!.session.part())
          .filter((site) => foreignIds.has(site.node.id))
          .map((site) => [site.node.id, JSON.stringify(site.node)])
      )
    ).toEqual(before);
    expect(
      groups(editor).some((group) => foreign.ownership.some((owned) => owned.groupId === group.id))
    ).toBe(true);
  } finally {
    editor.destroy();
  }
}, 30000);

for (const change of ['retain', 'split'] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`default GFM ${change} preserves native vertical geometry and ${decision}`, async () => {
      const editor = editor_();
      try {
        const original = await committed(editor),
          vertical =
            '**Alpha XYZ end**\n\n<table><tr><td>\n\nH1\n\n</td><td>\n\nH2\n\n</td></tr><tr><td rowspan="2">\n\nGroup\n\n</td><td>\n\nB\n\n</td></tr><tr><td>\n\nD\n\n</td></tr></table>\n\nTail';
        const first = await wave(editor, vertical),
          before = fingerprints(editor);
        let target = (await working(editor, false))
          .replace('Group', 'Final **group**')
          .replace('Alpha XYZ', 'Alpha Q');
        let expected = vertical.replace('Group', 'Final **group**').replace('Alpha XYZ', 'Alpha Q');
        if (change === 'split') {
          target = target.replace('|  | D |', '| New cell | D |');
          expected = target;
        }
        await wave(editor, target, first);
        expect(ast(await working(editor))).toBe(ast(expected));
        expect(ast(await committed(editor))).toBe(ast(original));
        const after = fingerprints(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprints(editor)).toEqual(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprints(editor)).toEqual(after);
        for (const group of [...groups(editor)]) {
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            group.id,
            groupToken(editor),
            groupToken(editor)
          );
          if (!result.success) throw new Error(result.error);
        }
        expect(ast(await working(editor))).toBe(ast(decision === 'accept' ? expected : original));
        const read = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!read.ok) throw new Error(read.reason);
        const fresh = propertyEditor(read.package);
        try {
          expect(ast(await working(fresh))).toBe(ast(decision === 'accept' ? expected : original));
        } finally {
          fresh.destroy();
        }
      } finally {
        editor.destroy();
      }
    }, 30000);

test('a host callback edit during geometry normalization refuses the stale target before replacement', async () => {
  const editor = editor_();
  try {
    const first = await wave(editor, htmlTarget),
      ids = groups(editor).map((group) => group.id);
    const raw = (await working(editor, false)).replace('**Alpha XYZ end**', '# Alpha XYZ end');
    let calls = 0;
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision: groupToken(editor),
        targetMarkdown: raw,
        ownership: first.ownership,
        contentProofs: first.contentProofs,
      },
      context,
      {
        target: {
          headingStyle: () => {
            calls++;
            const tail = editor.surface!.session.paragraphIds().at(-1)!;
            expect(
              editor.surface!.applyAutomationOps(
                () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' manual' }],
                { kind: 'body' },
                [],
                false,
                true,
                'after'
              ).committed
            ).toBe(true);
            return 'Heading1';
          },
        },
      }
    );
    expect(calls).toBe(1);
    expect(result).toMatchObject({ success: false, code: 'stale-revision' });
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    expect(ast(await working(editor))).toBe(ast(htmlTarget.replace('Tail', 'Tail manual')));
  } finally {
    editor.destroy();
  }
}, 30000);
