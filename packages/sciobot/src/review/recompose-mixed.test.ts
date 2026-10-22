import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groups, groupToken } from '../../test/support/review-groups';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
import { createEditPlan } from '../diff/plan';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals, type RecomposeOptions } from './recompose';
import {
  replaceOwnedProposals,
  type OwnedProposalReplacementResult,
} from './replace-owned-proposals';

const context = { ...propertyContext, date: '2026-10-21T14:00:00.000Z' };
type Editor = ReturnType<typeof propertyEditor>;
type Success = Extract<OwnedProposalReplacementResult, { success: true }>;
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
function fixture(body = p('Alpha end'), options: Parameters<typeof docxFixture>[1] = {}) {
  const read = readOoxmlPackage(docxFixture(body, options));
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
const pkg = (editor: Editor) => editor.surface!.session.currentPackage();
const fingerprints = (source: OoxmlPackage) =>
  [...source.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]);
const request = (editor: Editor, targetMarkdown: string, previous?: Success) => ({
  revision: groupToken(editor),
  targetMarkdown,
  ownership: previous?.ownership ?? [],
  contentProofs: previous?.contentProofs ?? [],
});
function ast(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
}
async function views(editor: Editor) {
  const original = await readCommittedView(editor.surface!.session, groupToken(editor));
  const proposed = await readWorkingView(editor.surface!.session, groupToken(editor));
  if (!original.success || !proposed.success) throw new Error('Missing mixed review views.');
  return [original.view.projection.text, proposed.view.projection.text];
}
async function expectViews(editor: Editor, original: string, proposed: string) {
  expect((await views(editor)).map(ast)).toEqual([original, proposed].map(ast));
}
async function wave(
  editor: Editor,
  target: string,
  previous?: Success,
  options?: RecomposeOptions
) {
  let publications = 0;
  const unsubscribe = editor.surface!.session.subscribe(() => publications++);
  try {
    const result = await recomposeOwnedProposals(
      editor,
      request(editor, target, previous),
      context,
      options
    );
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    expect(result.committed).toBe(true);
    expect(publications).toBe(1);
    return result;
  } finally {
    unsubscribe();
  }
}
async function reopen(editor: Editor, original: string, proposed: string) {
  const read = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!read.ok) throw new Error(read.reason);
  const fresh = propertyEditor(read.package);
  try {
    await expectViews(fresh, original, proposed);
    expect(groups(fresh).map((group) => group.id)).toEqual(groups(editor).map((group) => group.id));
  } finally {
    fresh.destroy();
  }
}

for (const target of [
  '**Alpha XYZ end**',
  '*A end*',
  '**Alpha much longer ending**',
  '~~A~~ end',
]) {
  test(`mixed post-text marks map exact native slices: ${target}`, async () => {
    const editor = propertyEditor(fixture());
    try {
      const before = fingerprints(pkg(editor));
      const first = await wave(editor, target);
      await expectViews(editor, 'Alpha end', target);
      const after = fingerprints(pkg(editor));
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprints(pkg(editor))).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(fingerprints(pkg(editor))).toEqual(after);
      const second = await wave(editor, '*Alpha Q end*', first);
      await expectViews(editor, 'Alpha end', '*Alpha Q end*');
      expect(second.ownership.length).toBeGreaterThan(first.ownership.length);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprints(pkg(editor))).toEqual(after);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      await reopen(editor, 'Alpha end', '*Alpha Q end*');
    } finally {
      editor.destroy();
    }
  }, 30_000);
}

test('mixed sibling waves retain paragraph identities and owned marks, links, and code', async () => {
  const editor = propertyEditor(fixture(p('Alpha end') + p('Keep') + p('Tail')));
  try {
    const original = 'Alpha end\n\nKeep\n\nTail';
    const ids = [...editor.surface!.session.paragraphIds()];
    const one =
      '**Alpha XYZ end**\n\nKeep\n\nNew **bold** and `code` [link](https://example.com/one)\n\nTail';
    const first = await wave(editor, one);
    await expectViews(editor, original, one);
    for (const id of ids) expect(editor.surface!.session.paragraphIds()).toContain(id);
    const afterFirst = fingerprints(pkg(editor));
    const two = '*Alpha Q end*\n\nKeep\n\nFinal *italic* [link](https://example.com/two)\n\nTail';
    const second = await wave(editor, two, first);
    await expectViews(editor, original, two);
    for (const id of ids) expect(editor.surface!.session.paragraphIds()).toContain(id);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(afterFirst);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    await reopen(editor, original, two);
    await wave(editor, 'Alpha final end\n\nKeep\n\nTail', second);
    await expectViews(editor, original, 'Alpha final end\n\nKeep\n\nTail');
    expect(editor.surface!.session.paragraphIds()).toEqual(ids);
  } finally {
    editor.destroy();
  }
}, 30_000);

test('mixed text and tracked heading conversion preserve paragraph properties and native decisions', async () => {
  const styles = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`;
  const editor = propertyEditor(fixture(p('Alpha end') + p('Keep'), { styles }));
  const options: RecomposeOptions = {
    target: { headingStyle: (depth) => (depth === 1 ? 'Heading1' : null) },
  };
  try {
    const first = await wave(editor, '# Alpha XYZ end\n\nKeep', undefined, options);
    await expectViews(editor, 'Alpha end\n\nKeep', '# Alpha XYZ end\n\nKeep');
    const afterFirst = fingerprints(pkg(editor));
    await wave(editor, '# Alpha Q end\n\nKeep', first, options);
    await expectViews(editor, 'Alpha end\n\nKeep', '# Alpha Q end\n\nKeep');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(afterFirst);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    await reopen(editor, 'Alpha end\n\nKeep', '# Alpha Q end\n\nKeep');
    for (const group of [...groups(editor)]) {
      const result = acceptReviewGroup(editor, group.id, groupToken(editor), context);
      if (!result.success) throw new Error(result.error);
    }
    await expectViews(editor, '# Alpha Q end\n\nKeep', '# Alpha Q end\n\nKeep');
    expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(0);
  } finally {
    editor.destroy();
  }
}, 30_000);

test('mixed retained cell text and marks use one native history unit with body text', async () => {
  const table =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>' +
    `<w:tc><w:tcPr/>${p('Cell one')}</w:tc><w:tc><w:tcPr/>${p('Cell two')}</w:tc></w:tr></w:tbl>`;
  const editor = propertyEditor(fixture(p('Alpha end') + table + p('Tail')));
  try {
    const original = (await views(editor))[0]!;
    const target = '**Alpha XYZ end**\n\n| Cell **ONE** | Cell two |\n| --- | --- |\n\nTail';
    const first = await wave(editor, target);
    await expectViews(editor, original, target);
    const afterFirst = fingerprints(pkg(editor));
    const second = '*Alpha Q end*\n\n| Cell *new* | Cell two |\n| --- | --- |\n\nTail';
    await wave(editor, second, first);
    await expectViews(editor, original, second);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprints(pkg(editor))).toEqual(afterFirst);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    await reopen(editor, original, second);
  } finally {
    editor.destroy();
  }
}, 30_000);

test('mixed hyperlink retarget and body text retain original targets through rejection', async () => {
  const editor = propertyEditor(
    fixture(
      p('Alpha end') +
        '<w:p><w:hyperlink r:id="old"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>',
      {
        relationships:
          '<Relationship Id="old" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/old" TargetMode="External"/>',
      }
    )
  );
  try {
    const original = 'Alpha end\n\n[Link](https://example.com/old)';
    const first = await wave(editor, '**Alpha XYZ end**\n\n[Link](https://example.com/new)');
    const target = '*Alpha Q end*\n\n[Link](https://example.com/final)';
    await wave(editor, target, first);
    await expectViews(editor, original, target);
    await reopen(editor, original, target);
    for (const group of [...groups(editor)]) {
      const result = rejectReviewGroup(editor, group.id, groupToken(editor), context);
      if (!result.success) throw new Error(result.error);
    }
    await expectViews(editor, original, original);
  } finally {
    editor.destroy();
  }
}, 30_000);

test('disabled formatting refuses the whole wave and preserves selected proposals and redo', async () => {
  const editor = propertyEditor(fixture(p('Alpha end') + p('Tail')));
  try {
    const first = await wave(editor, 'Alpha XYZ end\n\nTail');
    const session = editor.surface!.session;
    expect(
      editor.surface!.applyAutomationOps(
        () => [
          { op: 'insertText', paragraphId: session.paragraphIds()[1]!, offset: 4, text: ' redo' },
        ],
        { kind: 'body' },
        undefined,
        true,
        true
      ).committed
    ).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    editor.setRevisionMarkup({ trackFormatting: false });
    const before = pkg(editor),
      revision = groupToken(editor),
      prior = groups(editor);
    let publications = 0;
    const unsubscribe = session.subscribe(() => publications++);
    try {
      expect(
        await recomposeOwnedProposals(
          editor,
          request(editor, '**Alpha Q end**\n\nTail', first),
          context
        )
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(pkg(editor)).toBe(before);
      expect(groupToken(editor)).toEqual(revision);
      expect(groups(editor)).toEqual(prior);
      expect(publications).toBe(0);
    } finally {
      unsubscribe();
    }
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), session.paragraphIds()[1]!)).toBe('Tail redo');
  } finally {
    editor.destroy();
  }
}, 30_000);

test('replacement inherits the direct source run across mixed native styles', async () => {
  const editor = propertyEditor(
    fixture(
      '<w:p>' +
        '<w:r><w:t>a</w:t></w:r>' +
        '<w:r><w:rPr><w:b/><w:rFonts w:ascii="Arial"/><w:lang w:val="en-US"/></w:rPr><w:t>BC</w:t></w:r>' +
        '<w:r><w:rPr><w:i/><w:rFonts w:ascii="Times New Roman"/><w:lang w:val="cs-CZ"/></w:rPr><w:t>DE</w:t></w:r>' +
        '<w:r><w:t>f</w:t></w:r></w:p>'
    )
  );
  try {
    const original = (await views(editor))[0]!;
    const first = await wave(editor, 'a**X**f');
    await expectViews(editor, original, 'a**X**f');
    const second = await wave(editor, 'a*Y*f', first);
    await expectViews(editor, original, 'a*Y*f');
    expect(second.ownership.length).toBeGreaterThan(first.ownership.length);
    await reopen(editor, original, 'a*Y*f');
  } finally {
    editor.destroy();
  }
}, 30_000);

for (const kind of ['run', 'paragraph'] as const) {
  test(`manual current ${kind} properties invalidate owned format evidence before replacement`, async () => {
    const styles = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`;
    const editor = propertyEditor(fixture(p('Alpha end'), { styles }));
    const options: RecomposeOptions = { target: { headingStyle: () => 'Heading1' } };
    try {
      const first = await wave(
        editor,
        kind === 'run' ? '**Alpha end**' : '# Alpha end',
        undefined,
        options
      );
      const session = editor.surface!.session;
      const paragraphId = session.paragraphIds()[0]!;
      expect(
        editor.surface!.applyAutomationOps(
          () => [
            kind === 'run'
              ? {
                  op: 'setRunProperties',
                  paragraphId,
                  start: 0,
                  end: 9,
                  properties: [{ localName: 'i' }],
                }
              : {
                  op: 'setParagraphProperties',
                  paragraphId,
                  properties: [
                    { localName: 'pStyle', attributes: { val: 'Heading1' } },
                    { localName: 'jc', attributes: { val: 'center' } },
                  ],
                },
          ],
          { kind: 'body' },
          [],
          false,
          true
        ).committed
      ).toBe(true);
      const before = pkg(editor),
        revision = groupToken(editor),
        prior = groups(editor);
      let publications = 0;
      const unsubscribe = session.subscribe(() => publications++);
      try {
        const result = await recomposeOwnedProposals(
          editor,
          request(editor, kind === 'run' ? '*Alpha Q end*' : '# Alpha Q end', first),
          context,
          options
        );
        expect(result).toMatchObject({ success: false, code: 'ambiguous-target' });
        expect(pkg(editor)).toBe(before);
        expect(groupToken(editor)).toEqual(revision);
        expect(groups(editor)).toEqual(prior);
        expect(publications).toBe(0);
      } finally {
        unsubscribe();
      }
    } finally {
      editor.destroy();
    }
  }, 30_000);
}

test('an unselected same-author format keeps its current properties and complete group membership', async () => {
  const editor = propertyEditor(fixture(p('Alpha end') + p('Tail')));
  try {
    const foreign = await wave(editor, '**Alpha end**\n\nTail');
    const before = pkg(editor),
      revision = groupToken(editor),
      prior = groups(editor);
    const result = await recomposeOwnedProposals(
      editor,
      request(editor, '*Alpha end*\n\nTail'),
      context
    );
    expect(result).toMatchObject({ success: false, code: 'ambiguous-target' });
    expect(pkg(editor)).toBe(before);
    expect(groupToken(editor)).toEqual(revision);
    expect(groups(editor)).toEqual(prior);
    const own = await wave(editor, '**Alpha end**\n\nTail changed\n\nNew paragraph');
    const oldIds = new Set(foreign.ownership.map((entry) => entry.groupId));
    expect(own.ownership.some((entry) => oldIds.has(entry.groupId))).toBe(false);
    await expectViews(
      editor,
      'Alpha end\n\nTail',
      '**Alpha end**\n\nTail changed\n\nNew paragraph'
    );
  } finally {
    editor.destroy();
  }
}, 30_000);

for (const kind of ['inserted', 'deleted'] as const) {
  test(`manual paragraph properties inside an owned ${kind} sibling refuse before replacement`, async () => {
    const original = kind === 'inserted' ? p('Alpha end') : p('Alpha end') + p('Owned paragraph');
    const firstTarget = kind === 'inserted' ? 'Alpha end\n\nOwned paragraph' : 'Alpha end';
    const editor = propertyEditor(fixture(original));
    try {
      const first = await wave(editor, firstTarget);
      const session = editor.surface!.session;
      const paragraphId = session
        .paragraphIds()
        .find((id) => paragraphTextOf(session.part(), id) === 'Owned paragraph')!;
      expect(paragraphId).toBeDefined();
      expect(
        editor.surface!.applyAutomationOps(
          () => [
            {
              op: 'setParagraphProperties',
              paragraphId,
              properties: [{ localName: 'jc', attributes: { val: 'center' } }],
            },
          ],
          { kind: 'body' },
          [],
          false,
          true
        ).committed
      ).toBe(true);
      const target = kind === 'inserted' ? 'Alpha end\n\nReplacement paragraph' : 'Alpha changed';
      const refuse = async (current: Editor) => {
        const before = pkg(current),
          revision = groupToken(current),
          prior = groups(current);
        let publications = 0;
        const stop = current.surface!.session.subscribe(() => publications++);
        try {
          expect(
            await recomposeOwnedProposals(current, request(current, target, first), context)
          ).toMatchObject({ success: false, code: 'ambiguous-target' });
          expect(pkg(current)).toBe(before);
          expect(groupToken(current)).toEqual(revision);
          expect(groups(current)).toEqual(prior);
          expect(publications).toBe(0);
        } finally {
          stop();
        }
      };
      await refuse(editor);
      const read = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!read.ok) throw new Error(read.reason);
      const fresh = propertyEditor(read.package);
      try {
        await refuse(fresh);
      } finally {
        fresh.destroy();
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      await wave(editor, target, first);
    } finally {
      editor.destroy();
    }
  }, 30_000);
}

for (const destination of [
  { kind: 'external', authored: 'https://example.com/new', markdown: 'https://example.com/new' },
  { kind: 'internal', authored: 'Goal', markdown: '#Goal' },
] as const) {
  for (const order of ['tooltip-first', 'target-first'] as const) {
    test(`mixed hyperlink ${destination.kind} changes retain ${order} deltas through history and rejection`, async () => {
      const editor = propertyEditor(
        fixture(
          '<w:p><w:hyperlink r:id="old" w:tooltip="old title"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>',
          {
            relationships:
              '<Relationship Id="old" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/old" TargetMode="External"/>',
          }
        )
      );
      try {
        const session = editor.surface!.session;
        const paragraph = findNode(session.part(), session.paragraphIds()[0]!)!;
        if (paragraph.kind !== 'paragraph') throw new Error('Missing hyperlink paragraph.');
        const link = paragraph.children.find((node) => node.kind === 'hyperlink')!;
        const deltas = [
          {
            name: 'tooltip',
            before: { present: true, value: 'old title' },
            after: { present: true, value: 'new title' },
          },
          {
            name: 'target',
            before: {
              present: true,
              value: { kind: 'external', authored: 'https://example.com/old' },
            },
            after: {
              present: true,
              value: { kind: destination.kind, authored: destination.authored },
            },
          },
        ] as const;
        const planned = createEditPlan(groupToken(editor), [
          {
            kind: 'format',
            partName: session.part().name,
            target: { kind: 'hyperlink', hyperlinkId: link.id },
            changes: order === 'tooltip-first' ? deltas : [...deltas].reverse(),
          },
        ]);
        if (!planned.success) throw new Error(planned.error);
        const before = fingerprints(pkg(editor));
        let publications = 0;
        const stop = session.subscribe(() => publications++);
        try {
          expect(replaceOwnedProposals(editor, [], planned.plan, context)).toMatchObject({
            success: true,
            committed: true,
          });
          expect(publications).toBe(1);
        } finally {
          stop();
        }
        const original = '[Link](https://example.com/old "old title")';
        const target = `[Link](${destination.markdown} "new title")`;
        await expectViews(editor, original, target);
        const after = fingerprints(pkg(editor));
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprints(pkg(editor))).toEqual(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprints(pkg(editor))).toEqual(after);
        await reopen(editor, original, target);
        for (const group of [...groups(editor)]) {
          const result = rejectReviewGroup(editor, group.id, groupToken(editor), context);
          if (!result.success) throw new Error(result.error);
        }
        await expectViews(editor, original, original);
      } finally {
        editor.destroy();
      }
    }, 30_000);
  }
}
