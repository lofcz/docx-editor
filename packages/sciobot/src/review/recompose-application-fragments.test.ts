import { expect, test } from 'bun:test';
import {
  applicationFragmentTarget as target,
  applicationFragmentOptions as options,
} from '../../test/support/application-fragments';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import { MDX_BLOCK } from '../projection/mdx';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { readReviewGroups } from './groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(pkg.mainDocumentPart)!);
};
const key = (text: string) => {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
};
for (const decision of ['accept', 'reject'] as const)
  test(`AI application fragments retain nested payloads, two waves, ${decision}, exact Word output, and history`, async () => {
    const input = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
    if (!input.ok) throw Error(input.reason);
    const editor = propertyEditor(input.package);
    try {
      const baseline = editor.surface!.session.currentPackage();
      const baselineView = await readWorkingView(editor.surface!.session, propertyToken(editor));
      if (!baselineView.success) throw Error(baselineView.error);
      const directPlan = planTargetDocument(
        {
          pkg: baseline,
          revision: propertyToken(editor),
          projection: baselineView.view.projection,
        },
        target('Second'),
        options('second')
      );
      if (!directPlan.success) throw Error(directPlan.error);
      // Removing a born subtree retains its numbering definitions. Model that resource history
      // independently with the direct compiler; the second list gets the next native instance.
      const firstDirectPlan = planTargetDocument(
        {
          pkg: baseline,
          revision: propertyToken(editor),
          projection: baselineView.view.projection,
        },
        target('First'),
        options('first')
      );
      if (!firstDirectPlan.success) throw Error(firstDirectPlan.error);
      const firstDirect = previewDirectPlan(baseline, firstDirectPlan.plan);
      if (!firstDirect.success) throw Error(firstDirect.error);
      const direct = previewDirectPlan(
        withPart(firstDirect.pkg, baseline.parts.get(baseline.mainDocumentPart)!),
        directPlan.plan
      );
      if (!direct.success) throw Error(direct.error);
      const first = await recomposeOwnedProposals(
        editor,
        { revision: propertyToken(editor), targetMarkdown: target('First'), ownership: [] },
        propertyContext,
        { target: options('first'), payloads: { presence: { proposalId: 'birth-first' } } }
      );
      if (!first.success) throw Error(first.error);
      const pendingFirst = word(editor.surface!.session.currentPackage());
      const second = await recomposeOwnedProposals(
        editor,
        {
          revision: propertyToken(editor),
          targetMarkdown: target('Second'),
          ownership: first.ownership,
          contentProofs: first.contentProofs,
        },
        propertyContext,
        { target: options('second'), payloads: { presence: { proposalId: 'birth-second' } } }
      );
      if (!second.success) throw Error(second.error);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(4);
      const original = await readCommittedView(editor.surface!.session, propertyToken(editor));
      const working = await readWorkingView(editor.surface!.session, propertyToken(editor));
      if (!original.success || !working.success) throw Error('Missing review view');
      expect(key(original.view.projection.text)).toBe(key('Outside'));
      expect(key(working.view.projection.text)).toBe(key(target('Second')));
      const opaque = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(
        'second-4'
      )!.payload;
      expect(opaque.type).toBe(MDX_BLOCK);
      expect(opaque.attributes).toEqual({ host: 'react' });
      expect(opaque.extension).toEqual({ keep: [null, 2] });
      expect(JSON.stringify(opaque.mdast)).toContain('BudgetCard');
      expect(JSON.stringify(opaque.mdast)).toContain('rate: 2');
      const pendingSecond = word(editor.surface!.session.currentPackage());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(word(editor.surface!.session.currentPackage())).toBe(pendingFirst);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(word(editor.surface!.session.currentPackage())).toBe(pendingSecond);
      const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
        editor,
        propertyToken(editor),
        propertyContext
      );
      if (!decided.success) throw Error(decided.error);
      const expected = word(decision === 'accept' ? direct.pkg : baseline);
      expect(word(editor.surface!.session.currentPackage())).toBe(expected);
      const groups = readReviewGroups(editor.surface!.session.currentPackage());
      expect(groups.success && groups.groups.length).toBe(0);
      const payloads = readPayloadStore(editor.surface!.session.currentPackage());
      expect(payloads.status === 'ready' && payloads.entries.size).toBe(
        decision === 'accept' ? 4 : 0
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(word(editor.surface!.session.currentPackage())).toBe(pendingSecond);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(word(editor.surface!.session.currentPackage())).toBe(expected);
      const saved = readOoxmlPackage(writeOoxmlPackage(editor.surface!.session.currentPackage()));
      if (!saved.ok) throw Error(saved.reason);
      const reopened = propertyEditor(saved.package);
      try {
        expect(word(reopened.surface!.session.currentPackage())).toBe(expected);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  }, 30000);

test('missing presence identity refuses before native publication and leaves history unchanged', async () => {
  const input = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
  if (!input.ok) throw Error(input.reason);
  const editor = propertyEditor(input.package);
  try {
    const before = editor.surface!.session.currentPackage();
    const result = await recomposeOwnedProposals(
      editor,
      { revision: propertyToken(editor), targetMarkdown: target('First'), ownership: [] },
      propertyContext,
      { target: options('missing') }
    );
    expect(result.success).toBe(false);
    if (!result.success)
      expect(result.error).toBe(
        'Inserted application blocks need an explicit presence proposal identity.'
      );
    expect(editor.surface!.session.currentPackage()).toBe(before);
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('unsupported inline source insertion refuses without native publication', async () => {
  const input = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
  if (!input.ok) throw Error(input.reason);
  const editor = propertyEditor(input.package);
  try {
    const before = editor.surface!.session.currentPackage();
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision: propertyToken(editor),
        targetMarkdown:
          'Outside\n\n<activity name="Source" duration="5">\n<source title="Reference" url="https://example.com" type="verified" />\n</activity>',
        ownership: [],
      },
      propertyContext,
      { target: options('inline'), payloads: { presence: { proposalId: 'inline-birth' } } }
    );
    expect(result.success).toBe(false);
    if (!result.success)
      expect(result.error).toBe(
        'Inserted inline atoms need native asset and payload reconciliation.'
      );
    expect(editor.surface!.session.currentPackage()).toBe(before);
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

for (const invalid of ['id', 'ast'] as const)
  test(`invalid host React payload ${invalid} refuses before publication`, async () => {
    const input = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
    if (!input.ok) throw Error(input.reason);
    const editor = propertyEditor(input.package);
    try {
      const before = editor.surface!.session.currentPackage();
      const configured = options('invalid');
      const result = await recomposeOwnedProposals(
        editor,
        { revision: propertyToken(editor), targetMarkdown: target('First'), ownership: [] },
        propertyContext,
        {
          target: {
            ...configured,
            newBlockPayload: (ast, id) => {
              const authored = configured.newBlockPayload!(ast, id);
              return !authored
                ? null
                : {
                    ...authored,
                    payload: {
                      ...authored.payload,
                      ...(invalid === 'id'
                        ? { id: 'wrong-id' }
                        : {
                            mdast: {
                              type: 'mdxJsxFlowElement',
                              name: 'Other',
                              attributes: [],
                              children: [],
                            },
                          }),
                    },
                  };
            },
          },
          payloads: { presence: { proposalId: 'invalid-birth' } },
        }
      );
      expect(result.success).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });
