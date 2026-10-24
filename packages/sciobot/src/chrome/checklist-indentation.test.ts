import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  paragraphTextOf,
  findNode,
  replaceNode,
  applyTreeOp,
} from '@docx-editor.dev/core/store';
import fixtures from '../__tests__/fixtures/native-commands.json';
import { docxFixture } from '../../test/support/document';
import { checklistParagraph, listPPr } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { acceptReviewGroup } from '../review/accept-group';
import { readEditorRevision } from '../projection/revision-cache';
import { indentManualChecklist, manualListState, manualListMarker, setManualList } from './lists';
const context = { documentId: 'captured-checklist-indent', epoch: 'mount' };
for (const prior of [false, true])
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`captured checklist indentation preserves disc, dormant checked state, nesting and exact Word histories: prior=${prior}/${decision}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(
          prior
            ? checklistParagraph(false).replace('Task', 'Alpha beta')
            : '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>'
        ),
        author: 'User',
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      try {
        const surface = editor.surface!,
          id = surface.session.paragraphIds()[0]!;
        surface.setSelection({
          anchor: { paragraphId: id, offset: prior ? 1 : 0 },
          head: { paragraphId: id, offset: prior ? 6 : 5 },
        });
        const original = canonicalOoxmlFingerprint(surface.session.part());
        const capture = fixtures.cases.find(
          (item) => item.id === `list.todo.${prior ? 'off' : 'on'}`
        )!;
        const expected = capture.after.value[0] as Record<string, unknown>;
        if (decision !== 'direct')
          expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        const result = indentManualChecklist(editor, context);
        if (!result.success) throw new Error(result.error);
        const state = manualListState(editor, context, id)!;
        expect(state.checkbox).toBeNull();
        expect(state.dormantChecked).toBe(expected.checked as boolean);
        expect(state.numbering!.level + 1).toBe(expected.indent as number);
        expect(manualListMarker(editor, id)!.text).toBe('•');
        expect(expected.listStyleType).toBe('disc');
        const pending = canonicalOoxmlFingerprint(surface.session.part());
        const oracle = readOoxmlPackage(
          docxFixture(
            '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:pPr>' +
              listPPr(state.numbering!.numId, state.numbering!.level) +
              (prior ? '<w:customKeep w:val="yes"/>' : '') +
              '<scioList:state xmlns:scioList="urn:sciobot:docx:lists:1" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="scioList" mc:PreserveElements="scioList:state" checked="false"/></w:pPr><w:r><w:t>Alpha beta</w:t></w:r></w:p>'
          )
        );
        if (!oracle.ok) throw new Error(oracle.reason);
        const oraclePart = oracle.package.parts.get(oracle.package.mainDocumentPart)!;
        const expectedParagraph = oraclePart.root.children
          .flatMap((node) => (node.kind === 'textValue' ? [] : node.children))
          .find((node) => node.kind === 'paragraph')!;
        const sourceParagraph = findNode(surface.session.part(), id)!;
        if (expectedParagraph.kind !== 'paragraph' || sourceParagraph.kind !== 'paragraph')
          throw new Error('Missing oracle paragraph.');
        const withIdentity = replaceNode(oraclePart, expectedParagraph.id, {
          ...expectedParagraph,
          attributes: sourceParagraph.attributes,
        });
        if (!withIdentity.ok) throw new Error('Invalid oracle identity.');
        const accepted = canonicalOoxmlFingerprint(withIdentity.part);
        if (decision === 'direct') {
          expect(surface.state().selection).toEqual({
            anchor: { paragraphId: id, offset: 0 },
            head: { paragraphId: id, offset: 5 },
          });
          expect(pending).toEqual(accepted);
        } else {
          const groups = readReviewGroups(surface.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          // Word records an empty pPr snapshot when the original holder was absent.
          // Keep the complete native inverse as the rejection oracle, including that holder.
          const inverse = applyTreeOp(surface.session.part(), {
            op: 'rejectAllRevisions',
            siteNodeIds: groups.groups[0]!.native.flatMap((member) => member.siteNodeIds),
          });
          if (!inverse.ok) throw new Error(inverse.reason);
          const rejected = canonicalOoxmlFingerprint(inverse.part);
          if (prior) expect(rejected).toEqual(original);
          const decided = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            groups.groups[0]!.id,
            readEditorRevision(editor, context.documentId, context.epoch)!,
            context
          );
          if (!decided.success) throw new Error(decided.error);
          expect(canonicalOoxmlFingerprint(surface.session.part())).toEqual(
            decision === 'accept' ? accepted : rejected
          );
          if (decision === 'reject')
            expect(manualListState(editor, context, id)!.dormantChecked).toBeUndefined();
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(surface.session.part())).toEqual(pending);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        }
        const final = canonicalOoxmlFingerprint(surface.session.part());
        const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!reopened.ok) throw new Error(reopened.reason);
        expect(
          canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
        ).toEqual(final);
        expect(paragraphTextOf(surface.session.part(), id)).toBe(
          prior && decision === 'reject' ? '☐Alpha beta' : 'Alpha beta'
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        if (decision === 'direct')
          expect(canonicalOoxmlFingerprint(surface.session.part())).toEqual(original);
      } finally {
        editor.destroy();
      }
    });

test('removing list membership clears dormant check state and restores it through undo', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Item</w:t></w:r></w:p>'),
  });
  try {
    expect(indentManualChecklist(editor, context).success).toBe(true);
    const pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
    expect(setManualList(editor, 'none', context).success).toBe(true);
    expect(manualListState(editor, context)!.dormantChecked).toBeUndefined();
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(pending);
  } finally {
    editor.destroy();
  }
});

for (const extra of [
  '<s:state checked="maybe"/>',
  '<s:state checked="false"/><s:state checked="true"/>',
  '<s:state checked="false"><s:future/></s:state>',
  '<s:state checked="false" future="1"/>',
])
  test(`invalid dormant state refuses before native mutation: ${extra}`, () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        '<w:p><w:pPr xmlns:s="urn:sciobot:docx:lists:1">' +
          extra +
          '</w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
      ),
    });
    try {
      const surface = editor.surface!,
        version = surface.session.packageRevision(),
        before = canonicalOoxmlFingerprint(surface.session.part());
      expect(indentManualChecklist(editor, context).success).toBe(false);
      expect(surface.session.packageRevision()).toBe(version);
      expect(canonicalOoxmlFingerprint(surface.session.part())).toEqual(before);
    } finally {
      editor.destroy();
    }
  });

test('inactive imported dormant metadata stays lossless without creating an unreviewable list edit', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    author: 'User',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    document: docxFixture(
      '<w:p><w:pPr><s:state xmlns:s="urn:sciobot:docx:lists:1" checked="false"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
    ),
  });
  try {
    const surface = editor.surface!,
      before = canonicalOoxmlFingerprint(surface.session.part());
    expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
    expect(manualListState(editor, context)!.dormantChecked).toBeUndefined();
    const result = setManualList(editor, 'none', context);
    if (!result.success) throw new Error(result.error);
    expect(result.committed).toBe(false);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toEqual(before);
    const groups = readReviewGroups(surface.session.currentPackage());
    if (!groups.success) throw new Error(groups.error);
    expect(groups.groups).toHaveLength(0);
  } finally {
    editor.destroy();
  }
});
