import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  findNode,
  replaceNode,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import { readCommittedView } from '../review/committed-view';
import { readReviewGroups } from '../review/groups';
import { recomposeOwnedProposals } from '../review/recompose';
import { acceptAllReviewGroups } from '../review/accept-all';
import { rejectAllReviewGroups } from '../review/reject-all';

const context = {
  documentId: 'inline-birth',
  epoch: 'load',
  author: 'AI',
  date: '2026-10-24T12:00:00.000Z',
};
const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
// Native Word IDs are minted from local node allocation. Bind only this fresh identity into
// the independently lowered expected tree; every source identity and other native field stays exact.
function expectedBirth(direct: OoxmlPackage, pending: OoxmlPackage): OoxmlPackage {
  const actualBlock = buildBlockIndex(pending).byId.get('born-atom')!;
  const expectedBlock = buildBlockIndex(direct).byId.get('born-atom')!;
  const part = direct.parts.get(direct.mainDocumentPart)!;
  const control = findNode(part, expectedBlock.controlId)!;
  const actual = findNode(pending.parts.get(pending.mainDocumentPart)!, actualBlock.controlId)!;
  if (control.kind !== 'contentControl' || actual.kind !== 'contentControl')
    throw Error('Missing native control');
  const leaf = (node: typeof control) =>
    node.children
      .find((child) => child.kind === 'contentControlProperties')
      ?.children.find(
        (child) =>
          child.kind === 'generic' &&
          child.namespaceUri === WML_NAMESPACE_URI &&
          child.localName === 'id'
      );
  const allocated = leaf(actual),
    expected = leaf(control);
  if (!allocated || !expected || allocated.kind !== 'generic' || expected.kind !== 'generic')
    throw Error('Missing Word control ID');
  const value = allocated.attributes.find(
    (attr) => attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === 'val'
  )!.value;
  expect(value).toMatch(/^[1-9][0-9]*$/);
  const bound = replaceNode(part, expected.id, {
    ...expected,
    attributes: expected.attributes.map((attr) =>
      attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === 'val'
        ? { ...attr, value }
        : attr
    ),
  });
  if (!bound.ok) throw Error('Unable to bind allocated Word identity');
  return withPart(direct, bound.part);
}
const atoms = {
  mention: '<scio-mention key="user" value="Ada" />',
  source: '<source title="Paper" url="https://source.test" />',
  math: '$x+y$',
};
for (const kind of ['mention', 'source', 'math'] as const)
  for (const mark of [
    'plain',
    'bold',
    'code',
    'link',
    'link-start',
    'link-middle',
    'link-end',
    'bookmark',
    'retained-start',
    'retained-middle',
    'retained-end',
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`AI tracked retained inline birth: ${kind}/${mark}/${decision}`, async () => {
        const loaded = readOoxmlPackage(
          docxFixture(
            mark.startsWith('retained-')
              ? '<w:p><w:hyperlink w:anchor="target" w:tooltip="Keep" x:private="yes" xmlns:x="urn:test"><w:r><w:t xml:space="preserve">Lead.  Tail.</w:t></w:r></w:hyperlink></w:p>'
              : '<w:p><w:r><w:t xml:space="preserve">Lead.  Tail.</w:t></w:r></w:p>'
          )
        );
        if (!loaded.ok) throw Error(loaded.reason);
        const editor = createDocxEditor({
          container: document.createElement('div'),
          document: writeOoxmlPackage(loaded.package),
          modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        });
        const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
        try {
          const surface = editor.surface!,
            before = surface.session.currentPackage();
          const initial = await readWorkingView(surface.session, token());
          if (!initial.success) throw Error(initial.error);
          const authored =
            mark === 'bold'
              ? '**' + atoms[kind] + '**'
              : mark === 'code'
                ? '<code>' + atoms[kind] + '</code>'
                : atoms[kind];
          let target =
            mark === 'link'
              ? 'Lead. [' + authored + '](https://target.test "Keep") Tail.'
              : mark === 'link-start'
                ? 'Lead. [' + authored + ' Tail.](https://target.test "Keep")'
                : mark === 'link-middle'
                  ? '[Lead. ' + authored + ' Tail.](https://target.test "Keep")'
                  : mark === 'link-end'
                    ? '[Lead. ' + authored + '](https://target.test "Keep") Tail.'
                    : 'Lead. ' + authored + ' Tail.';
          if (mark === 'bookmark') target = 'Lead. [' + authored + '](#target "Keep") Tail.';
          if (mark === 'retained-start') target = '[' + authored + 'Lead.  Tail.](#target "Keep")';
          if (mark === 'retained-middle') target = '[Lead. ' + authored + ' Tail.](#target "Keep")';
          if (mark === 'retained-end') target = '[Lead.  Tail.' + authored + '](#target "Keep")';
          const options = { allocateBlockId: () => 'born-atom' };
          const plan = planTargetDocument(
            { pkg: before, revision: token(), projection: initial.view.projection },
            target,
            options
          );
          if (!plan.success) throw Error(plan.error);
          const direct = previewDirectPlan(before, plan.plan);
          if (!direct.success) throw Error(direct.error);
          const result = await recomposeOwnedProposals(
            editor,
            { revision: token(), targetMarkdown: target, ownership: [] },
            context,
            { target: options, payloads: { presence: { proposalId: 'birth-presence' } } }
          );
          if (!result.success) throw Error(result.error);
          expect(result.committed).toBe(true);
          const accepted = expectedBirth(direct.pkg, surface.session.currentPackage());
          const pending = native(surface.session.currentPackage());
          expect(buildBlockIndex(surface.session.currentPackage()).byId.size).toBe(1);
          const groups = readReviewGroups(surface.session.currentPackage());
          if (!groups.success) throw Error(groups.error);
          const presence = groups.groups.find((group) => group.properties.length)!;
          expect(presence.native).toHaveLength(1);
          const committed = await readCommittedView(surface.session, token());
          expect(committed.success && committed.view.projection.text).toBe(
            initial.view.projection.text
          );
          const working = await readWorkingView(surface.session, token());
          const directEditor = createDocxEditor({
            container: document.createElement('div'),
            document: writeOoxmlPackage(direct.pkg),
          });
          try {
            const expectedView = await readWorkingView(
              directEditor.surface!.session,
              readEditorRevision(directEditor, context.documentId, context.epoch)!
            );
            expect(working.success && working.view.projection.text).toBe(
              expectedView.success && expectedView.view.projection.text
            );
          } finally {
            directEditor.destroy();
          }
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(native(before));
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(pending);
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            editor,
            token(),
            context
          );
          if (!decided.success) throw Error(decided.error);
          expect(native(surface.session.currentPackage())).toBe(
            native(decision === 'accept' ? accepted : before)
          );
          const payloads = readPayloadStore(surface.session.currentPackage());
          expect(payloads.status === 'ready' && payloads.entries.size).toBe(
            decision === 'accept' ? 1 : 0
          );
          const decidedWord = native(surface.session.currentPackage());
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(pending);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(decidedWord);
          const saved = readOoxmlPackage(writeOoxmlPackage(surface.session.currentPackage()));
          if (!saved.ok) throw Error(saved.reason);
          expect(native(saved.package)).toBe(native(decision === 'accept' ? accepted : before));
        } finally {
          editor.destroy();
        }
      });

for (const linked of [false, true])
  test(`a later AI wave replaces locked inline births without retaining old native receipts or payloads: ${linked}`, async () => {
    const loaded = readOoxmlPackage(
      docxFixture('<w:p><w:r><w:t xml:space="preserve">Lead.  Tail.</w:t></w:r></w:p>')
    );
    if (!loaded.ok) throw Error(loaded.reason);
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(loaded.package),
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
    try {
      const surface = editor.surface!,
        before = surface.session.currentPackage();
      const first = await recomposeOwnedProposals(
        editor,
        {
          revision: token(),
          targetMarkdown: linked
            ? '[Lead. ' + atoms.mention + ' Tail.](https://target.test)'
            : 'Lead. ' + atoms.mention + ' Tail.',
          ownership: [],
        },
        context,
        {
          target: { allocateBlockId: () => 'first-atom' },
          payloads: { presence: { proposalId: 'first-birth' } },
        }
      );
      if (!first.success) throw Error(first.error);
      const firstWord = native(surface.session.currentPackage());
      const second = await recomposeOwnedProposals(
        editor,
        {
          revision: token(),
          targetMarkdown: linked
            ? '[Lead. ' + atoms.source + ' Tail.](https://other.test)'
            : 'Lead. ' + atoms.source + ' Tail.',
          ownership: first.ownership,
          contentProofs: first.contentProofs,
        },
        context,
        {
          target: { allocateBlockId: () => 'second-atom' },
          payloads: { presence: { proposalId: 'second-birth' } },
        }
      );
      if (!second.success) throw Error(second.error);
      const secondWord = native(surface.session.currentPackage());
      expect([...buildBlockIndex(surface.session.currentPackage()).byId.keys()]).toEqual([
        'second-atom',
      ]);
      const groups = readReviewGroups(surface.session.currentPackage());
      if (!groups.success) throw Error(groups.error);
      expect(groups.groups).toHaveLength(linked ? 2 : 1);
      expect(groups.groups.flatMap((group) => group.properties)).toEqual([
        { proposalId: 'second-birth', blockId: 'second-atom' },
      ]);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(native(surface.session.currentPackage())).toBe(firstWord);
      expect([...buildBlockIndex(surface.session.currentPackage()).byId.keys()]).toEqual([
        'first-atom',
      ]);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(native(surface.session.currentPackage())).toBe(secondWord);
      const rejected = rejectAllReviewGroups(editor, token(), context);
      if (!rejected.success) throw Error(rejected.error);
      expect(native(surface.session.currentPackage())).toBe(native(before));
      expect(buildBlockIndex(surface.session.currentPackage()).byId.size).toBe(0);
    } finally {
      editor.destroy();
    }
  });
