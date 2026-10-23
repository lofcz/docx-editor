import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  propertyContainer,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { JsonObject } from '../data/json';
import { planTableFormat } from '../diff/table-format';
import { planTableMerge } from '../diff/table-merges';
import {
  propertyContext,
  propertyEditor,
  executeProperties,
} from '../../test/support/property-execution';
import { propertyRecordFixture } from '../../test/support/property-record';
import {
  groupFixture,
  groups,
  groupToken,
  proposeGroupText,
  groupParagraph,
  insertLeadingGroup,
} from '../../test/support/review-groups';
import {
  groupDecisionFixture,
  proposeThreeProperties,
  crossStoryGroupFixture,
} from '../../test/support/group-decisions';
import {
  atomicFixture,
  atomicVisual,
  atomicListening,
  proposedVisual,
  proposedListening,
  executeAtomic,
} from '../../test/support/atomic-execution';
import { cellReviewDocument, cellScopeFor } from '../../test/support/table-review';
import { reviewDocumentFixture } from '../../test/support/review-document';
import { docxFixture } from '../../test/support/document';
import oracle from '../__tests__/fixtures/domain-properties.json';
import { readReviewGroups, type NativeReviewGroup } from './groups';
import {
  planOwnedProposalRemoval,
  type OwnedProposalGroup,
  type OwnedProposalRemovalPlan,
} from './owned-proposal-plan';
import {
  withInitializedReviewGroupStore,
  withPayloadGroupDecisionRecord,
  withReviewGroupRecord,
} from './group-record-store';
import { readPropertyReviewStore } from './property-record-store';
import { acceptReviewGroup } from './accept-group';
import { trackTableFormatPlan } from './track-table-format';
import { trackTableMergePlan } from './track-table-merges';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';

const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
const word = (pkg: OoxmlPackage) => canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
function exportedWord(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  return word(read.package);
}
const grant = (group: NativeReviewGroup): OwnedProposalGroup => ({
  groupId: group.id,
  native: group.native.map((member) => ({
    partName: member.partName,
    addresses: [member.address],
  })),
  properties: group.properties.map((member) => ({ ...member })),
});
function rejectSites(planned: OwnedProposalRemovalPlan) {
  const op = planned.partOps[0]!.ops[0]!;
  if (op.op !== 'rejectAllRevisions') throw new Error('Expected a native revision inverse');
  return op.siteNodeIds;
}
function plan(editor: DocxEditorInstance, selection = groups(editor).map(grant)) {
  const result = planOwnedProposalRemoval(
    { pkg: source(editor), revision: groupToken(editor) },
    selection
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Apply intents to a detached package only. This is not the R25 live executor. */
function inverse(pkg: OoxmlPackage, planned: OwnedProposalRemovalPlan) {
  let next = pkg;
  for (const { partName, ops } of planned.partOps) {
    let part = next.parts.get(partName)!;
    for (const op of ops) {
      const result = applyTreeOp(part, op);
      if (!result.ok) throw new Error(result.reason);
      part = result.part;
    }
    next = withPart(next, part);
  }
  if (planned.propertyDecisions.length || planned.retainNativeGroups.length) {
    const initialized = withInitializedReviewGroupStore(next);
    if (!initialized.ok) throw new Error(initialized.reason);
    next = initialized.pkg;
  }
  let index = 0;
  for (const retained of planned.retainNativeGroups) {
    const written = withReviewGroupRecord(next, {
      schemaVersion: 1,
      id: `rg_plan_native_${index++}`,
      members: retained.members,
    });
    if (!written.ok) throw new Error(written.reason);
    next = written.pkg;
  }
  for (const decision of planned.propertyDecisions) {
    const written = withPayloadGroupDecisionRecord(next, {
      schemaVersion: 2,
      id: `rg_plan_property_${index++}`,
      ...decision,
    });
    if (!written.ok) throw new Error(written.reason);
    next = written.pkg;
  }
  return next;
}
function readGroups(pkg: OoxmlPackage) {
  const read = readReviewGroups(pkg);
  if (!read.success) throw new Error(read.error);
  return read.groups;
}
function payloads(pkg: OoxmlPackage) {
  const read = readPayloadStore(pkg);
  if (read.status !== 'ready') throw new Error('Missing payloads');
  return [...read.entries].map(([id, entry]) => [id, entry.raw, entry.label]);
}

test('selected nested native inverse retains sibling proposals, canonical IDs, untouched references, and live history', async () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    proposeGroupText(editor, 'A dog ran.', 'A dog slept.');
    proposeGroupText(editor, 'Keep tail', 'Keep tail!');
    const pending = groups(editor),
      selection = [grant(pending[1]!)],
      pkg = source(editor),
      bytes = writeOoxmlPackage(pkg),
      revision = groupToken(editor),
      history = [editor.surface!.session.canUndo(), editor.surface!.session.canRedo()],
      tailId = pending[2]!.anchor!.paragraphId!,
      tail = findNode(editor.surface!.session.part(), tailId),
      planned = plan(editor, selection);
    expect(planned.revision).toEqual(revision);
    expect(planned.partOps).toHaveLength(1);
    expect(planned.partOps[0]!.ops[0]!.op).toBe('rejectAllRevisions');
    const after = inverse(pkg, planned),
      part = after.parts.get(after.mainDocumentPart)!;
    expect(paragraphTextOf(part, pending[1]!.anchor!.paragraphId!)).toBe('A dog ran.');
    expect(findNode(part, tailId)).toBe(tail);
    expect(readGroups(after)).toEqual([pending[0]!, pending[2]!]);
    expect(payloads(after)).toEqual(payloads(pkg));
    expect(source(editor)).toBe(pkg);
    expect(writeOoxmlPackage(pkg)).toEqual(bytes);
    expect(groupToken(editor)).toEqual(revision);
    expect([editor.surface!.session.canUndo(), editor.surface!.session.canRedo()]).toEqual(history);
    expect(plan(editor, selection)).toEqual(planned);
    const reopened = readOoxmlPackage(writeOoxmlPackage(after));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(readGroups(reopened.package).map((group) => group.id)).toEqual([
      pending[0]!.id,
      pending[2]!.id,
    ]);
    expect(word(reopened.package)).toBe(word(after));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual([pending[0]!.id, pending[1]!.id]);
  } finally {
    editor.destroy();
  }
});

test('explicit selection, not identical attribution, owns native groups; output order follows the document', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    proposeGroupText(editor, 'A dog ran.', 'A dog slept.');
    proposeGroupText(editor, 'Keep tail', 'Keep tail!');
    const pending = groups(editor),
      pkg = source(editor);
    const planned = plan(editor, [grant(pending[2]!), grant(pending[0]!)]);
    expect(planned.selection.map((group) => group.groupId)).toEqual([
      pending[0]!.id,
      pending[2]!.id,
    ]);
    expect(readGroups(inverse(pkg, planned))).toEqual([pending[1]!]);
    expect(rejectSites(planned)).not.toContain(pending[1]!.native[0]!.siteNodeIds[0]!);
  } finally {
    editor.destroy();
  }
});

for (const metadata of [false, true])
  test(`one property inverse retains exact canonical payloads and immutable proposal JSON; metadata=${metadata}`, () => {
    const editor = propertyEditor(groupDecisionFixture());
    try {
      proposeThreeProperties(editor, metadata);
      const pending = groups(editor),
        pkg = source(editor),
        stored = payloads(pkg),
        records = readPropertyReviewStore(pkg),
        selected = pending[1]!,
        planned = plan(editor, [grant(selected)]);
      expect(planned.propertyDecisions).toEqual([
        { proposalId: 'three-blocks', blockId: 'middle', decision: 'reject' },
      ]);
      expect(planned.partOps).toHaveLength(metadata ? 0 : 1);
      const after = inverse(pkg, planned);
      expect(payloads(after)).toEqual(stored);
      expect(buildBlockIndex(after).byId.get('middle')!.payload.attributes).toEqual({
        name: 'Old middle',
        duration: 10,
        resource: 'keep',
      });
      expect(readGroups(after)).toEqual([pending[0]!, pending[2]!]);
      const unchanged = readPropertyReviewStore(after);
      expect(
        records.status === 'ready' &&
          unchanged.status === 'ready' &&
          [...unchanged.records.values()].map((row) => row.raw)
      ).toEqual(records.status === 'ready' && [...records.records.values()].map((row) => row.raw));
      expect(groups(editor)).toEqual(pending);
    } finally {
      editor.destroy();
    }
  });

for (const item of oracle.cases)
  test('owned inverse retains captured property presence: ' + item.name, async () => {
    let pkg = propertyRecordFixture();
    const owner = buildBlockIndex(pkg).byId.get('activity')!,
      written = withPayloadRecord(
        pkg,
        { ...owner.payload, attributes: item.before as JsonObject },
        'Original label'
      );
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
    const editor = propertyEditor(pkg);
    try {
      const original = await readCommittedView(editor.surface!.session, groupToken(editor));
      if (!original.success) throw new Error(original.error);
      expect(
        executeProperties(editor, {
          set: Object.fromEntries(
            item.keys
              .filter((key) => Object.hasOwn(item.after, key))
              .map((key) => [key, (item.after as JsonObject)[key]!])
          ),
          remove: item.keys.filter((key) => !Object.hasOwn(item.after, key)),
        }).success
      ).toBe(true);
      const pending = source(editor),
        after = inverse(pending, plan(editor));
      expect(buildBlockIndex(after).byId.get('activity')!.payload.attributes).toEqual(
        item.committed as JsonObject
      );
      expect(payloads(after)).toEqual(payloads(pending));
      expect(readGroups(after)).toEqual([]);
      const reopened = propertyEditor(after);
      try {
        const current = await readWorkingView(reopened.surface!.session, groupToken(reopened));
        if (!current.success) throw new Error(current.error);
        expect(current.view.projection.text).toBe(original.view.projection.text);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

test('atomic inverse retains complete original assets and leaves the other pending atomic group unchanged', () => {
  const editor = propertyEditor(atomicFixture());
  try {
    expect(executeAtomic(editor, [proposedVisual, proposedListening]).success).toBe(true);
    const pending = groups(editor),
      pkg = source(editor),
      planned = plan(editor, [grant(pending[0]!)]),
      after = inverse(pkg, planned);
    expect(buildBlockIndex(after).byId.get('visual')!.payload).toEqual(atomicVisual);
    expect(buildBlockIndex(after).byId.get('audio')!.payload).toEqual(atomicListening);
    expect(payloads(after)).toEqual(payloads(pkg));
    expect(readGroups(after)).toEqual([pending[1]!]);
  } finally {
    editor.destroy();
  }
});

test('native run and paragraph property inverses match explicit committed formatting and retain old run identities', () => {
  const read = readOoxmlPackage(reviewDocumentFixture());
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const session = editor.surface!.session,
      firstId = session.paragraphIds()[0]!,
      headingId = session.paragraphIds()[1]!,
      before = session.part(),
      first = findNode(before, firstId)!;
    if (first.kind === 'textValue') throw new Error('Missing paragraph');
    const run = first.children.find((node) => node.kind === 'run')!;
    const after = inverse(source(editor), plan(editor)),
      part = after.parts.get(after.mainDocumentPart)!,
      oldRun = findNode(part, run.id)!;
    expect(paragraphTextOf(part, firstId)).toBe('Kept Old');
    expect(collectRevisionSites(part)).toEqual([]);
    const rPr = propertyContainer(oldRun, 'runProperties', 'rPr');
    if (!rPr || rPr.kind === 'textValue') throw new Error('Missing old run properties');
    expect(rPr.children.map((node) => (node.kind === 'textValue' ? '' : node.localName))).toEqual([
      'i',
    ]);
    const pPr = propertyContainer(findNode(part, headingId)!, 'paragraphProperties', 'pPr');
    if (!pPr || pPr.kind === 'textValue') throw new Error('Missing old paragraph properties');
    const outline = pPr.children[0];
    expect(
      outline &&
        outline.kind !== 'textValue' &&
        outline.attributes.find((attr) => attr.localName === 'val')?.value
    ).toBe('1');
    expect(findNode(part, run.id)?.id).toBe(run.id);
    expect(collectRevisionSites(before)).toHaveLength(4);
    expect(plan(editor).retainNativeGroups.length).toBeGreaterThan(0);
  } finally {
    editor.destroy();
  }
});

test('one owned group can inverse body and header parts while retaining foreign review and untouched parts', () => {
  const editor = propertyEditor(crossStoryGroupFixture());
  try {
    const pending = groups(editor),
      own = pending.find((group) => group.id === 'g1:stored:rg_cross_story')!,
      pkg = source(editor),
      planned = plan(editor, [grant(own)]),
      after = inverse(pkg, planned);
    expect(planned.partOps.map((row) => row.partName)).toEqual([
      '/word/document.xml',
      '/word/header1.xml',
    ]);
    expect(collectRevisionSites(after.parts.get('/word/header1.xml')!)).toEqual([]);
    expect(readGroups(after)).toEqual(pending.filter((group) => group.id !== own.id));
    for (const [name, part] of pkg.parts)
      if (!planned.partOps.some((row) => row.partName === name))
        expect(after.parts.get(name)).toBe(part);
  } finally {
    editor.destroy();
  }
});

test('owned paragraph insertion inverse retains existing canonical paragraphs and opaque payloads', () => {
  const editor = propertyEditor(groupFixture());
  try {
    const original = source(editor),
      ids = editor.surface!.session.paragraphIds();
    insertLeadingGroup(editor);
    const after = inverse(source(editor), plan(editor)),
      part = after.parts.get(after.mainDocumentPart)!;
    expect(exportedWord(after)).toBe(exportedWord(original));
    for (const id of ids)
      expect(findNode(part, id)).toBe(findNode(editor.surface!.session.part(), id));
    expect(payloads(after)).toEqual(payloads(original));
  } finally {
    editor.destroy();
  }
});

for (const kind of ['property', 'merge', 'split'] as const)
  test(`owned native ${kind} inverse restores exact committed table properties and node identities`, () => {
    const loaded = readOoxmlPackage(cellReviewDocument());
    if (!loaded.ok) throw new Error(loaded.reason);
    const editor = propertyEditor(loaded.package);
    try {
      const original = source(editor),
        scope = cellScopeFor(editor),
        input = {
          revision: groupToken(editor),
          part: editor.surface!.session.part(),
          tableId: scope.tableId,
        };
      const proposed =
        kind === 'property'
          ? planTableFormat(input, {
              op: 'setTableCellVerticalAlignment',
              cellIds: [scope.cellId],
              alignment: 'center',
            })
          : planTableMerge(
              input,
              kind === 'split'
                ? { kind: 'split', cellId: scope.cellId, rows: 1, columns: 2 }
                : { kind: 'merge', cellIds: [cellScopeFor(editor, 'One').cellId, scope.cellId] }
            );
      if (!proposed.success) throw new Error(proposed.error);
      expect(
        (kind === 'property' ? trackTableFormatPlan : trackTableMergePlan)(
          editor,
          proposed.plan,
          propertyContext
        ).success
      ).toBe(true);
      const after = inverse(source(editor), plan(editor));
      expect(word(after)).toBe(word(original));
      expect(findNode(after.parts.get(after.mainDocumentPart)!, scope.tableId)?.id).toBe(
        scope.tableId
      );
      expect(readGroups(after)).toEqual([]);
    } finally {
      editor.destroy();
    }
  });

test('planning preserves a prior redo frame and does not change native review display mode', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    const paragraphId = groupParagraph(editor, 'Keep tail');
    editor.surface!.applyAutomationOps(() => [
      { op: 'insertText', paragraphId, offset: 9, text: ' teacher' },
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.surface!.session.canRedo()).toBe(true);
    const pkg = source(editor),
      snapshot = editor.snapshot(),
      token = groupToken(editor);
    plan(editor);
    plan(editor);
    expect(source(editor)).toBe(pkg);
    expect(editor.snapshot()).toBe(snapshot);
    expect(groupToken(editor)).toEqual(token);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(editor.surface!.session.part(), paragraphId)).toBe('Keep tail teacher');
  } finally {
    editor.destroy();
  }
});

test('ownership validation refuses missing, partial, extra, forged, duplicate, and resolved group receipts without writes', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    proposeGroupText(editor, 'A dog ran.', 'A dog slept.');
    const pending = groups(editor),
      own = grant(pending[0]!),
      other = grant(pending[1]!),
      pkg = source(editor),
      bytes = writeOoxmlPackage(pkg);
    const invalid: OwnedProposalGroup[][] = [
      [{ ...own, groupId: 'missing' }],
      [own, own],
      [{ ...own, native: [] }],
      [{ ...own, native: own.native.slice(1) }],
      [{ ...own, native: [...own.native, ...other.native] }],
      [
        {
          ...own,
          native: own.native.map((member) => ({
            ...member,
            addresses: member.addresses.map((revision) => ({ ...revision, author: 'Foreign' })),
          })),
        },
      ],
      [{ ...own, native: [...own.native, own.native[0]!] }],
      [{ ...own, properties: [{ proposalId: 'foreign', blockId: 'activity' }] }],
    ];
    for (const selection of invalid) {
      expect(
        planOwnedProposalRemoval({ pkg, revision: groupToken(editor) }, selection).success
      ).toBe(false);
      expect(source(editor)).toBe(pkg);
    }
    expect(writeOoxmlPackage(pkg)).toEqual(bytes);
    expect(
      acceptReviewGroup(editor, own.groupId, groupToken(editor), propertyContext).success
    ).toBe(true);
    const accepted = source(editor);
    expect(
      planOwnedProposalRemoval({ pkg: accepted, revision: groupToken(editor) }, [own])
    ).toMatchObject({ success: false, code: 'stale-revision' });
    expect(source(editor)).toBe(accepted);
  } finally {
    editor.destroy();
  }
});

test('ownership data is bounded, detached, immutable, and never invokes accessors or serialization hooks', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    const own = grant(groups(editor)[0]!),
      selection = [{ ...own, native: structuredClone([...own.native]) }],
      planned = plan(editor, selection);
    selection[0]!.native.length = 0;
    expect(planned.selection[0]!.native).toHaveLength(own.native.length);
    expect(Object.isFrozen(planned)).toBe(true);
    expect(Object.isFrozen(rejectSites(planned))).toBe(true);
    expect(Object.isFrozen(planned.selection[0]!.native[0]!.addresses[0])).toBe(true);
    let invoked = false;
    const accessor = Object.defineProperty({}, 'groupId', {
      enumerable: true,
      get() {
        invoked = true;
        throw new Error('must not run');
      },
    });
    const hook = {
      ...own,
      toJSON() {
        invoked = true;
        throw new Error('must not run');
      },
    };
    const cycle: unknown[] = [];
    cycle.push(cycle);
    for (const invalid of [
      [accessor],
      [hook],
      cycle,
      [null],
      [{ ...own, extra: true }],
      new Array(200_001),
      [{ ...own, groupId: 'x'.repeat(16_385) }],
    ])
      expect(
        planOwnedProposalRemoval(
          { pkg: source(editor), revision: groupToken(editor) },
          invalid as OwnedProposalGroup[]
        ).success
      ).toBe(false);
    expect(invoked).toBe(false);
    const empty = plan(editor, []);
    expect(empty.partOps).toEqual([]);
    expect(empty.propertyDecisions).toEqual([]);
    expect(empty.retainNativeGroups).toEqual([]);
    for (const revision of [
      { ...groupToken(editor), revision: -1 },
      { ...groupToken(editor), epoch: '' },
      { ...groupToken(editor), revision: NaN },
    ])
      expect(planOwnedProposalRemoval({ pkg: source(editor), revision }, [own]).success).toBe(
        false
      );
  } finally {
    editor.destroy();
  }
});

for (const relation of ['owned-outer', 'owned-inner'] as const)
  test(`native nested revision protection keeps the foreign revision intact: ${relation}`, () => {
    const outerAuthor = relation === 'owned-outer' ? 'ScioBot' : 'Teacher',
      innerAuthor = relation === 'owned-inner' ? 'ScioBot' : 'Teacher',
      loaded = readOoxmlPackage(
        docxFixture(
          `<w:p><w:ins w:id="1" w:author="${outerAuthor}"><w:r><w:t>Keep</w:t></w:r><w:ins w:id="2" w:author="${innerAuthor}"><w:r><w:t>Inner</w:t></w:r></w:ins></w:ins></w:p>`
        )
      );
    if (!loaded.ok) throw new Error(loaded.reason);
    const editor = propertyEditor(loaded.package);
    try {
      const pkg = source(editor),
        bytes = writeOoxmlPackage(pkg),
        pending = groups(editor),
        selected = pending.find((group) =>
          group.native.some((member) => member.address.author === 'ScioBot')
        )!;
      const result = planOwnedProposalRemoval({ pkg, revision: groupToken(editor) }, [
        grant(selected),
      ]);
      if (relation === 'owned-outer') expect(result.success).toBe(false);
      else {
        if (!result.success) throw new Error(result.error);
        const after = inverse(pkg, result.plan),
          sites = collectRevisionSites(after.parts.get(after.mainDocumentPart)!);
        expect(sites).toHaveLength(1);
        expect(sites[0]!.node.attributes.find((attr) => attr.localName === 'author')?.value).toBe(
          'Teacher'
        );
        expect(
          paragraphTextOf(
            after.parts.get(after.mainDocumentPart)!,
            editor.surface!.session.paragraphIds()[0]!
          )
        ).toBe('Keep');
      }
      expect(source(editor)).toBe(pkg);
      expect(writeOoxmlPackage(pkg)).toEqual(bytes);
    } finally {
      editor.destroy();
    }
  });

test('owned paragraph inverse preserves a foreign pending paragraph property revision and its native owner', () => {
  const own = `w:id="17" w:author="ScioBot" w:date="${propertyContext.date}"`;
  const loaded = readOoxmlPackage(
    docxFixture(
      `<w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:ins ${own}/></w:rPr><w:pPrChange w:id="22" w:author="Peer"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:ins ${own}><w:r><w:t>New</w:t></w:r></w:ins></w:p>` +
        '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Control</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
        '<w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const initialized = withInitializedReviewGroupStore(loaded.package);
  if (!initialized.ok) throw new Error(initialized.reason);
  const written = withReviewGroupRecord(initialized.pkg, {
    schemaVersion: 1,
    id: 'rg_owned_peer_property',
    members: [
      {
        partName: loaded.package.mainDocumentPart,
        revision: { id: '17', author: 'ScioBot', date: propertyContext.date },
      },
    ],
  });
  if (!written.ok) throw new Error(written.reason);
  const editor = propertyEditor(written.pkg);
  try {
    const before = source(editor),
      part = before.parts.get(before.mainDocumentPart)!,
      body = part.root.children.find((node) => node.kind === 'body') as OoxmlElement,
      target = body.children[0]!,
      tail = body.children.at(-1)!,
      peer = collectRevisionSites(part).find((site) => site.node.localName === 'pPrChange')!,
      current = groups(editor),
      owned = current.find((group) => group.id === 'g1:stored:rg_owned_peer_property')!,
      foreign = current.filter((group) => group !== owned),
      beforeWord = word(before);
    expect(foreign).toHaveLength(1);
    const planned = plan(editor, [grant(owned)]),
      after = inverse(before, planned),
      afterPart = after.parts.get(after.mainDocumentPart)!,
      remaining = readGroups(after);
    expect(planned.propertyDecisions).toEqual([]);
    expect(rejectSites(planned)).not.toContain(peer.node.id);
    expect(remaining.map((group) => group.id)).toEqual(foreign.map((group) => group.id));
    expect(remaining[0]!.native).toEqual(foreign[0]!.native);
    expect(remaining[0]!.anchor!.paragraphId).toBe(target.id);
    expect(findNode(afterPart, peer.node.id)).toBe(peer.node);
    expect(paragraphTextOf(afterPart, target.id)).toBe('');
    expect(findNode(afterPart, tail.id)).toBe(tail);
    expect(collectRevisionSites(afterPart).map((site) => site.node.id)).toEqual([peer.node.id]);
    expect(source(editor)).toBe(before);
    expect(word(before)).toBe(beforeWord);
    const reopened = readOoxmlPackage(writeOoxmlPackage(after));
    if (!reopened.ok) throw new Error(reopened.reason);
    const savedSites = collectRevisionSites(
      reopened.package.parts.get(reopened.package.mainDocumentPart)!
    );
    expect(savedSites).toHaveLength(1);
    expect(savedSites[0]!.node.localName).toBe('pPrChange');
    expect(
      savedSites[0]!.node.attributes.find((attribute) => attribute.localName === 'author')?.value
    ).toBe('Peer');
  } finally {
    editor.destroy();
  }
});
