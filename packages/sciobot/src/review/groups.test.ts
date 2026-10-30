import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  groupFixture,
  groupParagraph,
  groups,
  insertLeadingGroup,
  proposeGroupText,
} from '../../test/support/review-groups';
import {
  executeProperties,
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import {
  executeAtomic,
  proposedListening,
  proposedVisual,
} from '../../test/support/atomic-execution';
import { blockContext, blockEditor, manualBlockPlan } from '../../test/support/block-execution';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { planDomainProperties } from '../diff/domain-properties';
import { createEditPlan } from '../diff/plan';
import { sourceControl, sourceParagraph } from '../../test/support/target-source-fixture';
import { trackPropertyPlan } from './track-properties';
import { CELL_CONTEXT, cellReviewDocument } from '../../test/support/table-review';
import { readEditorRevision } from '../projection/revision-cache';
import { planTableRow } from '../diff/table-rows';
import { planTableColumn } from '../diff/table-columns';
import { planTableMerge } from '../diff/table-merges';
import { planTableFormat } from '../diff/table-format';
import { trackTableRowPlan } from './track-table-rows';
import { trackTableColumnPlan } from './track-table-columns';
import { trackTableMergePlan } from './track-table-merges';
import { trackTableFormatPlan } from './track-table-format';
import { readReviewGroups, type NativeReviewGroup } from './groups';
import { reviewGroupMarker, REVIEW_GROUP_NAMESPACE } from './group-markers';
import {
  readReviewGroupStore,
  withInitializedReviewGroupStore,
  withReviewGroupRecord,
} from './group-record-store';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { ReviewGroupRecord } from './group-record-schema';
import { trackSiblingPlan } from './track-blocks';
import { trackHyperlinkPlan } from './track-hyperlinks';
import { LINK_CONTEXT, linkDocument, linkPlan } from '../../test/support/link-review';

const ids = (editor: ReturnType<typeof propertyEditor>) => groups(editor).map((group) => group.id);
function decideNative(
  editor: ReturnType<typeof propertyEditor>,
  group: NativeReviewGroup,
  op: 'acceptRevision' | 'rejectRevision'
) {
  const result = editor.surface!.applyAutomationOps(
    () =>
      group.native.map((member) => ({
        op,
        revision: member.address,
        siteNodeIds: member.siteNodeIds,
      })),
    { kind: 'body' },
    undefined,
    true,
    true
  );
  if (result.rejected) throw new Error(String(result.reason));
}

test('nested replacements persist one group per pair; pure reads and save/reopen retain exact identities', async () => {
  const editor = propertyEditor(groupFixture());
  try {
    const session = editor.surface!.session;
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    proposeGroupText(editor, 'A dog ran.', 'A hen ran.');
    const first = groups(editor),
      before = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision(),
      canUndo = session.canUndo();
    expect(first).toHaveLength(2);
    expect(first.map((group) => group.native.length)).toEqual([2, 2]);
    expect(first.every((group) => group.properties.length === 0 && !group.readOnly)).toBe(true);
    expect(first[1]!.anchor!.path.length).toBeGreaterThan(first[0]!.anchor!.path.length);
    expect(first[0]!.anchor!.paragraphId).not.toBe(first[1]!.anchor!.paragraphId);
    expect(first.every((group) => !group.id.includes('#new:'))).toBe(true);
    expect(Object.isFrozen(first)).toBe(true);
    for (const group of first)
      for (const member of group.native)
        for (const nodeId of member.siteNodeIds) {
          const node = findNode(session.part(), nodeId);
          if (!node || node.kind === 'textValue') throw new Error('Missing native group site');
          expect(reviewGroupMarker(node)).toBeUndefined();
        }
    const stored = readReviewGroupStore(session.currentPackage());
    if (stored.status !== 'ready') throw new Error('Missing persisted groups');
    expect([...stored.records.keys()].map((id) => `g1:stored:${id}`).sort()).toEqual(
      first.map((group) => group.id).sort()
    );
    for (const group of first) {
      const record = stored.records.get(group.id.slice('g1:stored:'.length))!.decoded;
      if (record.status !== 'supported') throw new Error('Unsupported group');
      expect(record.value.members).toEqual(
        group.native.map((member) => ({ partName: member.partName, revision: member.address }))
      );
    }
    expect(ids(editor)).toEqual(first.map((group) => group.id));
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
    expect(session.packageRevision()).toBe(revision);
    expect(session.canUndo()).toBe(canUndo);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    const reopened = propertyEditor(saved.package);
    try {
      expect(ids(reopened)).toEqual(ids(editor));
      expect(groups(reopened).map((group) => group.native.map((member) => member.address))).toEqual(
        first.map((group) => group.native.map((member) => member.address))
      );
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('unrelated insertion, decisions, and history preserve the other nested group IDs and native content', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    proposeGroupText(editor, 'A dog ran.', 'A hen ran.');
    const tailId = groupParagraph(editor, 'Keep tail');
    const initial = groups(editor),
      originalIds = ids(editor),
      kept = initial[1]!.native
        .flatMap((member) => member.siteNodeIds)
        .map(
          (id) =>
            [id, canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), id)!)] as const
        );
    insertLeadingGroup(editor);
    expect(groups(editor)).toHaveLength(3);
    expect(ids(editor).slice(1)).toEqual(originalIds);
    expect(groups(editor)[1]!.anchor!.path).not.toEqual(initial[0]!.anchor!.path);
    decideNative(editor, groups(editor)[1]!, 'acceptRevision');
    expect(groups(editor)).toHaveLength(2);
    expect(ids(editor)).toContain(originalIds[1]!);
    for (const [id, value] of kept)
      expect(canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), id)!)).toBe(value);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(ids(editor).slice(1)).toEqual(originalIds);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(ids(editor)).toContain(originalIds[1]!);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(ids(editor)).toEqual(originalIds);
    expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Keep tail');
  } finally {
    editor.destroy();
  }
});

test('property and atomic records form one group per block without swallowing their phase or neighboring text groups', async () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    const textId = ids(editor)[0]!;
    const property = executeProperties(editor),
      atomic = executeAtomic(editor, [proposedVisual, proposedListening]);
    if (!property.success) throw new Error(property.error);
    if (!atomic.success) throw new Error(atomic.error);
    const result = groups(editor);
    expect(result).toHaveLength(4);
    expect(result[0]!.id).toBe(textId);
    expect(result.map((group) => group.properties.map((member) => member.blockId))).toEqual([
      [],
      ['activity'],
      ['visual'],
      ['audio'],
    ]);
    expect(
      result
        .flatMap((group) => group.native)
        .flatMap((member) => member.siteNodeIds)
        .sort()
    ).toEqual(
      collectRevisionSites(editor.surface!.session.part())
        .map((site) => site.node.id)
        .sort()
    );
    const phase = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('phase')!;
    expect(result.every((group) => group.anchor?.nodeId !== phase.controlId)).toBe(true);
    const reopened = propertyEditor(editor.surface!.session.currentPackage());
    try {
      expect(ids(reopened)).toEqual(ids(editor));
    } finally {
      reopened.destroy();
    }
    const previous = ids(editor);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(ids(editor)).toEqual(previous.slice(0, 2));
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(ids(editor)).toEqual(previous);
  } finally {
    editor.destroy();
  }
});

test('metadata-only property proposals retain a navigable stable block group', () => {
  const editor = propertyEditor(groupFixture());
  try {
    const result = executeProperties(
      editor,
      { set: { duration: 15 } },
      'Old activity — 10 minutes'
    );
    if (!result.success) throw new Error(result.error);
    expect(groups(editor)).toHaveLength(1);
    expect(groups(editor)[0]!.native).toEqual([]);
    expect(groups(editor)[0]!.properties).toEqual([
      { proposalId: 'property-proposal', blockId: 'activity' },
    ]);
    expect(groups(editor)[0]!.anchor?.nodeId).toBe(
      buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('activity')!.controlId
    );
  } finally {
    editor.destroy();
  }
});

test('whole paragraph replacements persist one group across every inserted and deleted paragraph', () => {
  const mounted = blockEditor(
    ['Kept', 'Old first', 'Old second', 'Tail']
      .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
      .join('')
  );
  try {
    const result = trackSiblingPlan(
      mounted.editor,
      manualBlockPlan(mounted.editor, 1, 2, '## New\n\nNew body'),
      blockContext
    );
    if (!result.success) throw new Error(result.error);
    const found = groups(mounted.editor);
    expect(found).toHaveLength(1);
    expect(found[0]!.native.map((member) => member.address)).toEqual([
      ...result.proposals[0]!.addresses,
    ]);
    // Each paragraph retains both its paragraph-mark revision and its content revision.
    expect(found[0]!.native.flatMap((member) => member.siteNodeIds)).toHaveLength(8);
    const reopened = propertyEditor(mounted.editor.surface!.session.currentPackage());
    try {
      expect(ids(reopened)).toEqual(ids(mounted.editor));
    } finally {
      reopened.destroy();
    }
    const pending = ids(mounted.editor);
    expect(mounted.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(mounted.editor)).toEqual([]);
    expect(mounted.editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(ids(mounted.editor)).toEqual(pending);
  } finally {
    mounted.destroy();
  }
});

test('namespace prefix collisions retain foreign bindings and unknown native attributes', () => {
  const zip = unzipSync(writeOoxmlPackage(groupFixture()));
  zip['word/document.xml'] = strToU8(
    strFromU8(zip['word/document.xml']!).replace(
      '<w:document ',
      '<w:document xmlns:scioReview="urn:foreign-review" xmlns:scioMc="urn:foreign-mc" '
    )
  );
  const loaded = readOoxmlPackage(zipSync(zip));
  if (!loaded.ok) throw new Error(loaded.reason);
  const editor = propertyEditor(loaded.package);
  try {
    const root = editor.surface!.session.part().root.namespaceBindings;
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    expect(editor.surface!.session.part().root.namespaceBindings).toEqual(root);
    for (const site of collectRevisionSites(editor.surface!.session.part())) {
      expect(
        site.node.attributes.find((attribute) => attribute.namespaceUri === REVIEW_GROUP_NAMESPACE)
      ).toBeUndefined();
    }
    const reopened = propertyEditor(editor.surface!.session.currentPackage());
    try {
      expect(ids(reopened)).toEqual(ids(editor));
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('ambiguous imported reused addresses remain separate, visible, stable, and read-only without writing metadata', () => {
  const loaded = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="4" w:author="Teacher"><w:r><w:t>First</w:t></w:r></w:ins><w:r><w:t> gap </w:t></w:r><w:ins w:id="4" w:author="Teacher"><w:r><w:t>Second</w:t></w:r></w:ins></w:p>'
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const before = writeOoxmlPackage(loaded.package),
    result = readReviewGroups(loaded.package);
  if (!result.success) throw new Error(result.error);
  expect(result.groups).toHaveLength(2);
  expect(result.groups.every((group) => group.readOnly)).toBe(true);
  expect(new Set(result.groups.map((group) => group.id)).size).toBe(2);
  expect(writeOoxmlPackage(loaded.package)).toEqual(before);
  const read = readOoxmlPackage(before);
  if (!read.ok) throw new Error(read.reason);
  const again = readReviewGroups(read.package);
  if (!again.success) throw new Error(again.error);
  expect(again.groups.map((group) => group.id)).toEqual(result.groups.map((group) => group.id));
});

test('retargeted hyperlinks persist one replacement group across the old link and its proposed copy', () => {
  const read = readOoxmlPackage(linkDocument());
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const result = trackHyperlinkPlan(
      editor,
      linkPlan(editor, { url: 'https://example.org/new' }),
      LINK_CONTEXT
    );
    if (!result.success) throw new Error(result.error);
    expect(groups(editor)).toHaveLength(1);
    expect(groups(editor)[0]!.native.map((member) => member.address)).toEqual([
      ...result.proposals[0]!.addresses,
    ]);
    const reopened = propertyEditor(editor.surface!.session.currentPackage());
    try {
      expect(ids(reopened)).toEqual(ids(editor));
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('a grapheme plan with two separated replacements keeps two independently paired groups', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'A dog ran.', 'A fox ran.');
    expect(groups(editor)).toHaveLength(2);
    expect(groups(editor).every((group) => group.native.length === 2)).toBe(true);
    expect(new Set(ids(editor)).size).toBe(2);
  } finally {
    editor.destroy();
  }
});

test('an unsupported persisted group marker refuses a new transaction without changing content or history', () => {
  const read = readOoxmlPackage(
    docxFixture(
      `<w:p><w:ins w:id="9" w:author="Teacher" xmlns:g="${REVIEW_GROUP_NAMESPACE}" g:group="g2:future"><w:r><w:t>Peer</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>The cat sat.</w:t></w:r></w:p>`
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const session = editor.surface!.session,
      before = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision();
    expect(readReviewGroups(session.currentPackage()).success).toBe(false);
    expect(() => proposeGroupText(editor, 'The cat sat.', 'The car sat.')).toThrow(
      'Unsupported or malformed'
    );
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
    expect(session.packageRevision()).toBe(revision);
    expect(session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

function storeGroup(pkg: OoxmlPackage, id: string, members: ReviewGroupRecord['members']) {
  const initialized = withInitializedReviewGroupStore(pkg);
  if (!initialized.ok) throw new Error(initialized.reason);
  const written = withReviewGroupRecord(initialized.pkg, { schemaVersion: 1, id, members });
  if (!written.ok) throw new Error(written.reason);
  return written.pkg;
}

test('separate group records preserve original native compatibility attributes and XML', () => {
  const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const read = readOoxmlPackage(
    docxFixture(
      `<w:p><w:ins w:id="8" w:author="Teacher" xmlns:mc="${mc}" xmlns:keep="urn:keep" mc:Ignorable="keep" mc:PreserveAttributes="keep:flag" keep:flag="retained"><w:r><w:t>Text</w:t></w:r></w:ins></w:p>`
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const partName = read.package.mainDocumentPart,
    before = canonicalOoxmlFingerprint(read.package.parts.get(partName)!),
    candidate = storeGroup(read.package, 'rg_original', [
      { partName, revision: { id: '8', author: 'Teacher' } },
    ]);
  expect(candidate.parts.get(partName)).toBe(read.package.parts.get(partName));
  const reopened = readOoxmlPackage(writeOoxmlPackage(candidate));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(canonicalOoxmlFingerprint(reopened.package.parts.get(partName)!)).toBe(before);
  const again = readReviewGroups(reopened.package);
  if (!again.success) throw new Error(again.error);
  expect(again.groups.map((group) => group.id)).toEqual(['g1:stored:rg_original']);
});

test('independent parent and child records preserve nested native identities and original XML through save/reopen', () => {
  const read = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t xml:space="preserve">Added </w:t></w:r><w:del w:id="2" w:author="Bob"><w:r><w:delText>struck</w:delText></w:r></w:del></w:ins></w:p>'
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const part = read.package.parts.get(read.package.mainDocumentPart)!,
    before = canonicalOoxmlFingerprint(part),
    candidate = storeGroup(
      storeGroup(read.package, 'rg_child', [
        { partName: part.name, revision: { id: '2', author: 'Bob' } },
      ]),
      'rg_parent',
      [{ partName: part.name, revision: { id: '1', author: 'Alice' } }]
    );
  expect(candidate.parts.get(part.name)).toBe(part);
  expect(paragraphTextOf(part, storyParagraphs(bodyStoryRoot(part)!)[0]!.id)).toBe('Added struck');
  const reopened = readOoxmlPackage(writeOoxmlPackage(candidate));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(canonicalOoxmlFingerprint(reopened.package.parts.get(part.name)!)).toBe(before);
  const result = readReviewGroups(reopened.package);
  if (!result.success) throw new Error(result.error);
  expect(result.groups.map((group) => group.id)).toEqual([
    'g1:stored:rg_parent',
    'g1:stored:rg_child',
  ]);
});

test('partly resolved stored membership stays visible and read-only; fully resolved membership stays inactive', () => {
  const read = readOoxmlPackage(
    docxFixture('<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>Present</w:t></w:r></w:ins></w:p>')
  );
  if (!read.ok) throw new Error(read.reason);
  const partName = read.package.mainDocumentPart,
    candidate = storeGroup(read.package, 'rg_partial', [
      { partName, revision: { id: '1', author: 'Alice' } },
      { partName, revision: { id: '2', author: 'Alice' } },
    ]),
    inactive = storeGroup(candidate, 'rg_inactive', [
      { partName, revision: { id: '3', author: 'Bob' } },
    ]),
    before = writeOoxmlPackage(inactive),
    result = readReviewGroups(inactive);
  if (!result.success) throw new Error(result.error);
  expect(result.groups).toHaveLength(1);
  expect(result.groups[0]).toMatchObject({ id: 'g1:stored:rg_partial', readOnly: true });
  expect(result.groups[0]!.native).toHaveLength(1);
  expect(writeOoxmlPackage(inactive)).toEqual(before);
  const store = readReviewGroupStore(inactive);
  expect(store.status === 'ready' && store.records.size).toBe(2);
});

test('conflicting stored membership refuses pure reads and a new transaction without writes', () => {
  const read = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>Peer</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>The cat sat.</w:t></w:r></w:p>'
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const member = {
      partName: read.package.mainDocumentPart,
      revision: { id: '1', author: 'Alice' },
    },
    candidate = storeGroup(storeGroup(read.package, 'rg_first', [member]), 'rg_second', [member]),
    editor = propertyEditor(candidate);
  try {
    const session = editor.surface!.session,
      before = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision();
    expect(readReviewGroups(candidate)).toMatchObject({
      success: false,
      error: 'A native revision site has conflicting group ownership.',
    });
    expect(() => proposeGroupText(editor, 'The cat sat.', 'The car sat.')).toThrow(
      'conflicting group ownership'
    );
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
    expect(session.packageRevision()).toBe(revision);
    expect(session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('one property proposal partitions nested blocks by their own fallback instead of their shared ancestor', () => {
  const label = (text: string) =>
    '<w:sdt><w:sdtPr><w:tag w:val="scio:label"/></w:sdtPr><w:sdtContent>' +
    sourceParagraph(text) +
    '</w:sdtContent></w:sdt>';
  const read = readOoxmlPackage(
    docxFixture(
      sourceControl(
        'outer',
        label('Outer') + sourceControl('inner', label('Inner') + sourceParagraph('Body'))
      )
    )
  );
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const [id, name] of [
    ['outer', 'Outer'],
    ['inner', 'Inner'],
  ] as const) {
    const written = withPayloadRecord(
      pkg,
      { id, type: 'lesson_activity', version: 1, attributes: { name, duration: 10 } },
      name
    );
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  const editor = propertyEditor(pkg);
  try {
    const revision = propertyToken(editor),
      operations = ['outer', 'inner'].flatMap((blockId) => {
        const result = planDomainProperties(
          { pkg: editor.surface!.session.currentPackage(), revision, blockId },
          { set: { duration: 15 } }
        );
        if (!result.success) throw new Error(result.error);
        return result.plan.operations;
      }),
      plan = createEditPlan(revision, operations);
    if (!plan.success) throw new Error(plan.error);
    const result = trackPropertyPlan(
      editor,
      plan.plan,
      [
        { blockId: 'outer', before: 'Outer', after: 'Outer — 15' },
        { blockId: 'inner', before: 'Inner', after: 'Inner — 15' },
      ],
      propertyContext
    );
    if (!result.success) throw new Error(result.error);
    expect(groups(editor)).toHaveLength(2);
    expect(groups(editor).map((group) => group.properties.map((member) => member.blockId))).toEqual(
      [['outer'], ['inner']]
    );
    expect(groups(editor).every((group) => group.native.length > 0 && !group.readOnly)).toBe(true);
    const members = groups(editor).map((group) =>
      group.native.flatMap((member) => member.siteNodeIds)
    );
    expect(members[0]!.some((nodeId) => members[1]!.includes(nodeId))).toBe(false);
    const reopened = propertyEditor(editor.surface!.session.currentPackage());
    try {
      expect(ids(reopened)).toEqual(ids(editor));
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

for (const operation of ['row', 'column', 'merge', 'format'] as const)
  test(`tracked table ${operation} persists its complete native group through save and history`, () => {
    const read = readOoxmlPackage(cellReviewDocument());
    if (!read.ok) throw new Error(read.reason);
    const editor = propertyEditor(read.package);
    try {
      const session = editor.surface!.session,
        part = session.part(),
        body = bodyStoryRoot(part)!,
        table = body.kind !== 'textValue' && body.children.find((node) => node.kind === 'table');
      if (!table || table.kind !== 'table') throw new Error('Missing group table');
      const row = table.children.find((node) => node.kind === 'tableRow')!,
        cells = row.children.filter((node) => node.kind === 'tableCell'),
        grid = table.children.find((node) => node.localName === 'tblGrid')!,
        column = grid.children.find(
          (node) => node.kind !== 'textValue' && node.localName === 'gridCol'
        );
      if (!column) throw new Error('Missing group grid');
      const source = {
          revision: readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
          part,
          tableId: table.id,
        },
        plan =
          operation === 'row'
            ? planTableRow(source, {
                kind: 'insert',
                rowId: row.id,
                where: 'below',
                values: ['New A', 'New B'],
              })
            : operation === 'column'
              ? planTableColumn(source, { kind: 'insert', gridColumnId: column.id, where: 'right' })
              : operation === 'merge'
                ? planTableMerge(source, { kind: 'merge', cellIds: cells.map((node) => node.id) })
                : planTableFormat(source, {
                    op: 'setTableCellFill',
                    cellIds: [cells[0]!.id],
                    color: { kind: 'hex', value: '00FF00' },
                  });
      if (!plan.success) throw new Error(plan.error);
      const result =
        operation === 'row'
          ? trackTableRowPlan(editor, plan.plan, CELL_CONTEXT)
          : operation === 'column'
            ? trackTableColumnPlan(editor, plan.plan, CELL_CONTEXT)
            : operation === 'merge'
              ? trackTableMergePlan(editor, plan.plan, CELL_CONTEXT)
              : trackTableFormatPlan(editor, plan.plan, CELL_CONTEXT);
      if (!result.success) throw new Error(result.error);
      const pending = ids(editor);
      expect(pending).toHaveLength(1);
      expect(groups(editor)[0]!.native.map((member) => member.address)).toEqual([
        ...result.proposals[0]!.addresses,
      ]);
      const reopened = propertyEditor(session.currentPackage());
      try {
        expect(ids(reopened)).toEqual(pending);
      } finally {
        reopened.destroy();
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor)).toEqual([]);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(ids(editor)).toEqual(pending);
    } finally {
      editor.destroy();
    }
  });

test('a later proposal reserves the inactive single-member group IDs and restores either group through history', () => {
  const editor = propertyEditor(groupFixture());
  try {
    insertLeadingGroup(editor);
    const pending = groups(editor)[0]!;
    expect(pending.native).toHaveLength(1);
    decideNative(editor, pending, 'acceptRevision');
    expect(groups(editor)).toEqual([]);
    insertLeadingGroup(editor);
    const later = groups(editor)[0]!;
    expect(later.id).not.toBe(pending.id);
    expect(later.native.map((member) => member.address.id)).not.toEqual(
      pending.native.map((member) => member.address.id)
    );
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(ids(editor)).toEqual([pending.id]);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(ids(editor)).toEqual([later.id]);
  } finally {
    editor.destroy();
  }
});
