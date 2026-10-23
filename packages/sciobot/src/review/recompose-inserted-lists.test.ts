import { expect, test } from 'bun:test';
import {
  readOoxmlPackage,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groupToken, groups } from '../../test/support/review-groups';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { recomposeOwnedProposals } from './recompose';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
const context = { ...propertyContext, date: '2026-10-23T00:00:00Z' };
for (const [name, inserted, nativeExpected] of [
  ['bullet', '- **One**\n- Two', '- **One**\n\n- Two'],
  ['ordered', '3. One\n4. Two', '3. One\n\n4. Two'],
  ['nested', '- One\n  - Nested\n- Two', '- One\n\n  - Nested\n\n- Two'],
  ['checklist', '- [ ] One\n- [x] Two', '- [ ] One\n\n- [x] Two'],
] as const)
  test(`later-wave coordinator inserts native ${name} paragraphs`, async () => {
    const read = readOoxmlPackage(
      docxFixture('<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>')
    );
    if (!read.ok) throw Error(read.reason);
    const editor = propertyEditor(read.package),
      golden = propertyEditor(read.package);
    const target = `Before\n\n${inserted}\n\nAfter`;
    try {
      const expected = `Before\n\n${nativeExpected}\n\nAfter`;
      const base = await readWorkingView(golden.surface!.session, groupToken(golden));
      if (!base.success) throw Error(base.error);
      const directPlan = planTargetDocument(
        {
          revision: groupToken(golden),
          pkg: golden.surface!.session.currentPackage(),
          projection: base.view.projection,
        },
        target
      );
      if (!directPlan.success) throw Error(directPlan.error);
      const direct = applyDirectPlan(golden, directPlan.plan, context);
      if (!direct.success) throw Error(direct.error);
      const canonical = await readWorkingView(golden.surface!.session, groupToken(golden));
      if (!canonical.success) throw Error(canonical.error);
      expect(canonical.view.projection.text.trimEnd()).toBe(expected);
      const unchanged = planTargetDocument(
        {
          revision: groupToken(golden),
          pkg: golden.surface!.session.currentPackage(),
          projection: canonical.view.projection,
        },
        target
      );
      if (!unchanged.success) throw Error(unchanged.error);
      expect(unchanged.plan.operations).toEqual([]);
      const result = await recomposeOwnedProposals(
        editor,
        { revision: groupToken(editor), targetMarkdown: target, ownership: [], contentProofs: [] },
        context
      );
      if (!result.success) throw Error(`${result.code}: ${result.error}`);
      const original = await readCommittedView(editor.surface!.session, groupToken(editor)),
        proposed = await readWorkingView(editor.surface!.session, groupToken(editor));
      if (!original.success || !proposed.success) throw Error('Missing views');
      expect(original.view.projection.text.trimEnd()).toBe('Before\n\nAfter');
      expect(proposed.view.projection.text.trimEnd()).toBe(expected);
    } finally {
      editor.destroy();
      golden.destroy();
    }
  });

function importedParagraphKey(editor: ReturnType<typeof propertyEditor>, id: string): string {
  const part = editor.surface!.session.part();
  const paragraph = findNode(part, id);
  if (paragraph?.kind !== 'paragraph') throw Error('The imported paragraph disappeared');
  return canonicalOoxmlFingerprint({ ...part, root: paragraph });
}

const laterCases = [
  [
    'bullet',
    '- **One**\n- Two',
    '- **One revised**\n- Two\n- Three',
    '- **One revised**\n\n- Two\n\n- Three',
  ],
  [
    'ordered',
    '3. One\n4. Two',
    '3. One revised\n4. Two\n5. Three',
    '3. One revised\n\n4. Two\n\n5. Three',
  ],
  [
    'nested',
    '- One\n  - Nested\n- Two',
    '- One revised\n  - Nested\n  - Another\n- Two',
    '- One revised\n\n  - Nested\n\n  - Another\n\n- Two',
  ],
  [
    'checklist',
    '- [ ] One\n- [x] Two',
    '- [x] One revised\n- [x] Two\n- [ ] Three',
    '- [x] One revised\n\n- [x] Two\n\n- [ ] Three',
  ],
] as const;
for (const [name, firstBody, nextBody, nativeNext] of laterCases)
  for (const action of ['accept', 'reject'] as const)
    test(`two ${name} insertion waves keep exact history, imported paragraphs, ${action}, and saved numbering`, async () => {
      const loaded = readOoxmlPackage(
        docxFixture('<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>')
      );
      if (!loaded.ok) throw Error(loaded.reason);
      const editor = propertyEditor(loaded.package);
      const target = (body: string) => `Before\n\n${body}\n\nAfter`;
      const body = () => canonicalOoxmlFingerprint(editor.surface!.session.part());
      const views = async () => {
        const a = await readCommittedView(editor.surface!.session, groupToken(editor)),
          b = await readWorkingView(editor.surface!.session, groupToken(editor));
        if (!a.success || !b.success) throw Error('Missing views');
        return [a.view.projection.text.trimEnd(), b.view.projection.text.trimEnd()];
      };
      try {
        const original = body();
        const imported = editor
          .surface!.session.paragraphIds()
          .map((id) => [id, importedParagraphKey(editor, id)] as const);
        const first = await recomposeOwnedProposals(
          editor,
          {
            revision: groupToken(editor),
            targetMarkdown: target(firstBody),
            ownership: [],
            contentProofs: [],
          },
          context
        );
        if (!first.success) throw Error(first.error);
        const prior = body();
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(body()).toBe(original);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(body()).toBe(prior);
        const second = await recomposeOwnedProposals(
          editor,
          {
            revision: groupToken(editor),
            targetMarkdown: target(nextBody),
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          context
        );
        if (!second.success) throw Error(second.error);
        expect(await views()).toEqual(['Before\n\nAfter', target(nativeNext)]);
        expect(groups(editor)).toHaveLength(1);
        expect(
          collectRevisionSites(editor.surface!.session.part()).every((site) => site.nesting === 0)
        ).toBe(true);
        const pending = body();
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(body()).toBe(prior);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(body()).toBe(pending);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(await views()).toEqual(
          action === 'accept'
            ? [target(nativeNext), target(nativeNext)]
            : ['Before\n\nAfter', 'Before\n\nAfter']
        );
        expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(0);
        for (const [id, expected] of imported)
          expect(importedParagraphKey(editor, id)).toBe(expected);
        const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!opened.ok) throw Error(opened.reason);
        const numbering = opened.package.parts.get('/word/numbering.xml');
        const ids = new Set(
          numbering?.root.children.flatMap((node) =>
            node.kind !== 'textValue' && node.localName === 'num'
              ? [node.attributes.find((a) => a.localName === 'numId')?.value]
              : []
          ) ?? []
        );
        const references = [
          ...serializeOoxmlPart(
            opened.package.parts.get(opened.package.mainDocumentPart)!
          ).matchAll(/<w:numId w:val="([0-9]+)"/g),
        ]
          .map((match) => match[1]!)
          .filter((id) => id !== '0');
        for (const id of references) expect(ids.has(id)).toBe(true);
        const reopened = propertyEditor(opened.package);
        try {
          const projected = await readWorkingView(reopened.surface!.session, groupToken(reopened));
          if (!projected.success) throw Error(projected.error);
          expect(projected.view.projection.text.trimEnd()).toBe(
            action === 'accept' ? target(nativeNext) : 'Before\n\nAfter'
          );
        } finally {
          reopened.destroy();
        }
      } finally {
        editor.destroy();
      }
    });
