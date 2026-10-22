import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  headlessViewOfStore,
  openHeadlessDocument,
  readOoxmlPackage,
  TreePackageStore,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import oracle from '../__tests__/fixtures/domain-properties.json';
import { docxFixture } from '../../test/support/document';
import { sourceControl, sourceParagraph } from '../../test/support/target-source-fixture';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload, JsonObject } from '../blocks/payload-schema';
import { planDomainProperties, type DomainPropertyTarget } from '../diff/domain-properties';
import { createEditPlan, type CanonicalEdit, type CanonicalEditPlan } from '../diff/plan';
import { markdownRangeAt } from '../projection/markdown';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import {
  createPendingPropertyProposal,
  projectPendingProperties,
  type PendingPropertyProposal,
} from './property-projection';

const revision = { documentId: 'properties', epoch: 'load-1', revision: 0 };
function fixture(attributes: JsonObject = { name: 'Old', duration: 10 }, type = 'lesson_activity') {
  const loaded = readOoxmlPackage(
    docxFixture(sourceControl('activity', sourceParagraph('Body')) + sourceParagraph('Outside'))
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  return store(loaded.package, {
    id: 'activity',
    type,
    version: 1,
    attributes,
    extension: { opaque: { token: 'keep' } },
  });
}
function store(pkg: OoxmlPackage, payload: BlockPayload) {
  const result = withPayloadRecord(pkg, payload, 'Original label');
  if (!result.ok) throw new Error(result.reason);
  return result.pkg;
}
const block = (pkg: OoxmlPackage) => buildBlockIndex(pkg).byId.get('activity')!;
function propose(pkg: OoxmlPackage, target: DomainPropertyTarget, id = 'proposal-1') {
  const result = planDomainProperties({ pkg, revision, blockId: 'activity' }, target);
  if (!result.success) throw new Error(result.error);
  return proposal(result.plan, id);
}
function proposal(plan: CanonicalEditPlan, id = 'proposal-1') {
  const result = createPendingPropertyProposal(id, plan);
  if (!result.success) throw new Error(result.error);
  return result.proposal;
}
function view(
  pkg: OoxmlPackage,
  pending: readonly PendingPropertyProposal[],
  kind: 'committed' | 'working'
) {
  const result = projectPendingProperties(pkg, revision, pending, kind);
  if (!result.success) throw new Error(result.error);
  return result.pkg;
}

test.each(oracle.cases)(
  'native property views match captured application presence: $name',
  (item) => {
    const pkg = fixture(item.before as JsonObject),
      pending = propose(pkg, {
        set: Object.fromEntries(
          item.keys
            .filter((key) => Object.hasOwn(item.after, key))
            .map((key) => [key, (item.after as JsonObject)[key]!])
        ),
        remove: item.keys.filter((key) => !Object.hasOwn(item.after, key)),
      }),
      bytes = writeOoxmlPackage(pkg),
      current = store(pkg, { ...block(pkg).payload, attributes: item.after as JsonObject });
    for (const source of [pkg, current]) {
      const committed = view(source, [pending], 'committed'),
        working = view(source, [pending], 'working');
      expect(block(committed).payload.attributes).toEqual(item.committed as JsonObject);
      expect(block(working).payload.attributes).toEqual(item.working as JsonObject);
      expect(committed.parts.get(pkg.mainDocumentPart)).toBe(
        source.parts.get(pkg.mainDocumentPart)
      );
      expect(working.parts.get(pkg.mainDocumentPart)).toBe(source.parts.get(pkg.mainDocumentPart));
      expect(block(committed).payload.extension).toEqual({ opaque: { token: 'keep' } });
      expect(block(working).controlId).toBe(block(source).controlId);
      for (const [name, part] of source.parts) {
        const records = readPayloadStore(source);
        if (records.status === 'ready' && name === records.part.partName) continue;
        expect(committed.parts.get(name)).toBe(part);
        expect(working.parts.get(name)).toBe(part);
      }
    }
    expect(writeOoxmlPackage(pkg)).toEqual(bytes);
  }
);

test('pending values are detached, deeply frozen, and preserve false, zero, null, and absence', () => {
  const pkg = fixture({ name: null, duration: 0 }),
    owned = propose(pkg, { set: { name: { nested: [false, 0, null] } }, remove: ['duration'] });
  const raw = JSON.parse(JSON.stringify(owned.plan)) as CanonicalEditPlan,
    copied = proposal(raw);
  (
    raw.operations[0] as unknown as { after: { value: { nested: unknown[] } } }
  ).after.value.nested.push('mutated');
  expect(Object.isFrozen(copied)).toBe(true);
  expect(Object.isFrozen(copied.plan.operations)).toBe(true);
  const working = view(pkg, [copied], 'working'),
    committed = view(working, [copied], 'committed');
  expect(block(working).payload.attributes).toEqual({ name: { nested: [false, 0, null] } });
  expect(Object.hasOwn(block(working).payload.attributes, 'duration')).toBe(false);
  expect(block(committed).payload.attributes).toEqual({ name: null, duration: 0 });
  expect(Object.hasOwn(block(committed).payload.attributes, 'name')).toBe(true);
});

test('continuous pending waves retain their earliest committed and latest working values from every canonical intermediate', () => {
  const pkg = fixture({ duration: 10, resource: 'untouched' }),
    first = propose(pkg, { set: { duration: 15 } }, 'first'),
    middle = view(pkg, [first], 'working'),
    second = propose(middle, { set: { duration: 20, name: 'Added' } }, 'second'),
    final = view(middle, [second], 'working'),
    peer = store(middle, {
      ...block(middle).payload,
      attributes: { duration: 15, resource: 'peer' },
    });
  for (const source of [pkg, middle, final]) {
    expect(block(view(source, [first, second], 'committed')).payload.attributes).toEqual({
      duration: 10,
      resource: 'untouched',
    });
    expect(block(view(source, [first, second], 'working')).payload.attributes).toEqual({
      duration: 20,
      name: 'Added',
      resource: 'untouched',
    });
  }
  expect(block(view(peer, [first, second], 'committed')).payload.attributes).toEqual({
    duration: 10,
    resource: 'peer',
  });
  expect(block(view(peer, [first, second], 'working')).payload.attributes).toEqual({
    duration: 20,
    name: 'Added',
    resource: 'peer',
  });
  expect(projectPendingProperties(pkg, revision, [second, first], 'working')).toMatchObject({
    success: false,
    code: 'ambiguous-target',
  });
});

test('no pending values and already materialized views preserve package identity', () => {
  const pkg = fixture(),
    pending = propose(pkg, { set: { duration: 15 } });
  expect(view(pkg, [], 'committed')).toBe(pkg);
  expect(view(pkg, [], 'working')).toBe(pkg);
  expect(view(pkg, [pending], 'committed')).toBe(pkg);
  const current = view(pkg, [pending], 'working');
  expect(view(current, [pending], 'working')).toBe(current);
  expect(readPayloadStore(current)).toMatchObject({ status: 'ready' });
  const records = readPayloadStore(current);
  if (records.status !== 'ready') throw new Error('Missing native records');
  expect(records.entries.get('activity')!.label).toBe('Original label');
});

test('disjoint blocks and properties preserve independent source identities', () => {
  const first = fixture(),
    initial = readOoxmlPackage(
      docxFixture(
        sourceControl('activity', sourceParagraph('One')) +
          sourceControl('other', sourceParagraph('Two'))
      )
    );
  if (!initial.ok) throw new Error(initial.reason);
  const pkg = store(initial.package, block(first).payload),
    two = store(pkg, {
      ...block(first).payload,
      id: 'other',
      attributes: { name: 'Other', duration: 20 },
    }),
    a = propose(two, { set: { duration: 15 } }),
    result = planDomainProperties(
      { pkg: two, revision, blockId: 'other' },
      { set: { name: 'Second' } }
    );
  if (!result.success) throw new Error(result.error);
  const b = proposal(result.plan, 'second'),
    current = view(two, [a, b], 'working');
  expect(
    [...buildBlockIndex(current).byId.values()].map((item) => item.payload.attributes)
  ).toEqual([
    { name: 'Old', duration: 15 },
    { name: 'Second', duration: 20 },
  ]);
  expect(block(view(current, [a, b], 'committed')).payload.attributes.duration).toBe(10);
});

test('invalid or conflicting proposals refuse atomically without overwriting changed peer properties', () => {
  const pkg = fixture(),
    pending = propose(pkg, { set: { duration: 15 } }),
    bytes = writeOoxmlPackage(pkg),
    peer = store(pkg, { ...block(pkg).payload, attributes: { name: 'Old', duration: 12 } });
  const mutate = (change: (plan: CanonicalEditPlan) => void) => {
    const raw = JSON.parse(JSON.stringify(pending)) as PendingPropertyProposal;
    change(raw.plan);
    return raw;
  };
  const cases = [
    [pending, pending],
    [
      mutate((plan) => {
        (plan.operations[0] as { controlId: string }).controlId = 'missing';
      }),
    ],
    [
      mutate((plan) => {
        (plan.revision as { documentId: string }).documentId = 'other';
      }),
    ],
    [
      mutate((plan) => {
        (plan.operations[0] as unknown as { path: string[] }).path = [];
      }),
    ],
    [
      mutate((plan) => {
        (plan.operations[0] as unknown as { path: string[] }).path = ['attributes', 'privateData'];
      }),
    ],
    [
      mutate((plan) => {
        (plan.operations as CanonicalEdit[]).push(plan.operations[0]!);
      }),
    ],
  ];
  for (const values of cases) {
    const result = projectPendingProperties(pkg, revision, values, 'working');
    expect(result.success).toBe(false);
    expect('pkg' in result).toBe(false);
  }
  expect(projectPendingProperties(peer, revision, [pending], 'committed')).toMatchObject({
    success: false,
    code: 'stale-revision',
  });
  expect(
    projectPendingProperties(
      pkg,
      revision,
      [{ ...pending, extra: 'unknown' } as PendingPropertyProposal],
      'working'
    ).success
  ).toBe(false);
  expect(writeOoxmlPackage(pkg)).toEqual(bytes);
});

test('a canceled structural identity change cannot pass as a property-only pending chain', () => {
  const pkg = fixture({ grade: 7 }, 'lesson_title'),
    owner = block(pkg),
    edit: CanonicalEdit = {
      kind: 'payload',
      partName: owner.partName,
      blockId: owner.id,
      controlId: owner.controlId,
      path: ['attributes', 'grade'],
      before: { present: true, value: 7 },
      after: { present: true, value: 9 },
    },
    result = createEditPlan(revision, [edit]),
    inverse = createEditPlan(revision, [{ ...edit, before: edit.after, after: edit.before }]);
  if (!result.success || !inverse.success) throw new Error('Invalid property plan');
  expect(
    projectPendingProperties(
      pkg,
      revision,
      [proposal(result.plan, 'a'), proposal(inverse.plan, 'b')],
      'working'
    )
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
});

test('getter, cycle, undefined, oversized, and mixed native edit inputs never become property proposals', () => {
  const pkg = fixture(),
    pending = propose(pkg, { set: { duration: 15 } });
  let calls = 0;
  const getter = {
    ...pending,
    get plan(): CanonicalEditPlan {
      calls++;
      throw new Error('executed');
    },
  };
  expect(projectPendingProperties(pkg, revision, [getter], 'working').success).toBe(false);
  expect(calls).toBe(0);
  const cycle: Record<string, unknown> = {};
  cycle.self = cycle;
  for (const value of [undefined, cycle, 'x'.repeat(2 * 1024 * 1024 + 1)])
    expect(
      createPendingPropertyProposal('bad', value as unknown as CanonicalEditPlan).success
    ).toBe(false);
  const mixed = createEditPlan(revision, [
    ...pending.plan.operations,
    {
      kind: 'text',
      partName: pkg.mainDocumentPart,
      paragraphId: 'p',
      start: 0,
      end: 0,
      before: '',
      after: 'New',
    },
  ]);
  if (!mixed.success) throw new Error(mixed.error);
  expect(createPendingPropertyProposal('mixed', mixed.plan)).toMatchObject({
    success: false,
    code: 'unsupported-operation',
  });
  expect(createPendingPropertyProposal('  ', pending.plan).success).toBe(false);
  expect(
    projectPendingProperties(
      pkg,
      revision,
      Array.from({ length: 1025 }, (_, index) => ({ ...pending, id: String(index) })),
      'working'
    ).success
  ).toBe(false);
});

test('native review reads compose text revisions and property presence without changing live state or history', async () => {
  const loaded = readOoxmlPackage(
    docxFixture(
      sourceControl(
        'activity',
        '<w:p><w:del w:id="1" w:author="Ada"><w:r><w:delText>Old</w:delText></w:r></w:del><w:ins w:id="2" w:author="Ada"><w:r><w:t>New</w:t></w:r></w:ins></w:p>'
      )
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const pkg = store(loaded.package, {
      id: 'activity',
      type: 'lesson_activity',
      version: 1,
      attributes: { name: 'Activity', duration: 10 },
    }),
    pending = propose(pkg, { set: { duration: 15, resource: 'worksheet' } }),
    current = view(pkg, [pending], 'working'),
    editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(current),
    });
  try {
    const session = editor.surface!.session,
      token = readEditorRevision(editor, revision.documentId, revision.epoch)!,
      bytes = writeOoxmlPackage(session.currentPackage()),
      fingerprint = canonicalOoxmlFingerprint(session.part()),
      history = [session.canUndo(), session.canRedo()],
      count = collectRevisionSites(session.part()).length,
      mode = editor.surface!.layout().displayMode;
    const [committed, working] = await Promise.all([
      readCommittedView(session, token, { pendingProperties: [pending] }),
      readWorkingView(session, token, { pendingProperties: [pending] }),
    ]);
    if (!committed.success) throw new Error(committed.error);
    if (!working.success) throw new Error(working.error);
    expect(committed.view.projection.text).toBe(
      '<activity name="Activity" duration="10">\n\nOld\n\n</activity>'
    );
    expect(working.view.projection.text).toBe(
      '<activity name="Activity" duration="15" resource="worksheet">\n\nNew\n\n</activity>'
    );
    const start = working.view.projection.text.indexOf('New');
    expect(markdownRangeAt(working.view.projection, start, start + 3)).toMatchObject({
      start: { paragraphId: session.paragraphIds()[0], offset: 3 },
      end: { paragraphId: session.paragraphIds()[0], offset: 6 },
    });
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(session.packageRevision()).toBe(token.revision);
    expect(collectRevisionSites(session.part())).toHaveLength(count);
    expect(editor.surface!.layout().displayMode).toBe(mode);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(fingerprint);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    const reopened = openHeadlessDocument(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const next = { ...token, epoch: 'reopened', revision: reopened.view.packageRevision() },
      restored = JSON.parse(JSON.stringify(pending)) as PendingPropertyProposal,
      read = await readCommittedView(reopened.view, next, { pendingProperties: [restored] });
    if (!read.success) throw new Error(read.error);
    expect(read.view.projection.text).toBe(committed.view.projection.text);
    expect(block(reopened.view.currentPackage()).payload.attributes).toEqual({
      name: 'Activity',
      duration: 15,
      resource: 'worksheet',
    });
    expect(collectRevisionSites(reopened.view.part())).toHaveLength(count);
  } finally {
    editor.destroy();
  }
});

test('property options snapshot before an asynchronous native view read', async () => {
  const pkg = fixture(),
    pending = propose(pkg, { set: { duration: 15 } }),
    native = headlessViewOfStore(new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!)),
    raw = JSON.parse(JSON.stringify(pending)) as PendingPropertyProposal,
    reading = readWorkingView(native, revision, { pendingProperties: [raw] });
  (raw.plan.operations[0] as { after: { value: number } }).after.value = 999;
  const result = await reading;
  if (!result.success) throw new Error(result.error);
  expect(result.view.projection.text).toContain('duration="15"');
  expect(result.view.projection.text).not.toContain('999');
  expect(block(pkg).payload.attributes.duration).toBe(10);
});

test('a racing native edit discards the property view and preserves user history', async () => {
  const pkg = fixture(),
    pending = propose(pkg, { set: { duration: 15 } }),
    editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(pkg),
    });
  try {
    const surface = editor.surface!,
      session = surface.session,
      token = readEditorRevision(editor, revision.documentId, revision.epoch)!,
      reading = readWorkingView(session, token, { pendingProperties: [pending] }),
      id = session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    surface.insertPlainText('User ');
    const bytes = writeOoxmlPackage(session.currentPackage());
    expect(await reading).toMatchObject({ success: false, code: 'stale-revision' });
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(block(session.currentPackage()).payload.attributes.duration).toBe(10);
    expect(session.canUndo()).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});
