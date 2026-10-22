import { expect, test } from 'bun:test';
import {
  collectRevisionSites,
  customXmlNodes,
  withCustomXmlNode,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  groupFixture,
  groupParagraph,
  groups,
  insertLeadingGroup,
  proposeGroupText,
} from '../../test/support/review-groups';
import { propertyEditor } from '../../test/support/property-execution';
import { reviewGroupRevisionReservations } from './group-allocation';
import {
  readReviewGroupStore,
  withInitializedReviewGroupStore,
  withReviewGroupRecord,
  withoutReviewGroupRecord,
} from './group-record-store';

function fixture() {
  const initialized = withInitializedReviewGroupStore(groupFixture());
  if (!initialized.ok) throw new Error(initialized.reason);
  const written = withReviewGroupRecord(initialized.pkg, {
    schemaVersion: 1,
    id: 'rg_retired',
    members: [{ partName: '/word/document.xml', revision: { id: '100', author: 'ScioBot' } }],
  });
  if (!written.ok) throw new Error(written.reason);
  return written.pkg;
}

test('the private source reserves active, inactive, and removed native IDs without modifying any part', () => {
  const pkg = fixture(),
    before = writeOoxmlPackage(pkg),
    read = readReviewGroupStore(pkg);
  if (read.status !== 'ready') throw new Error('Missing group store');
  expect(reviewGroupRevisionReservations(pkg)).toEqual([
    { partName: '/word/document.xml', ids: ['100'] },
  ]);
  const mounted = propertyEditor(pkg);
  try {
    expect(groups(mounted)).toEqual([]);
  } finally {
    mounted.destroy();
  }
  expect(writeOoxmlPackage(pkg)).toEqual(before);
  const removed = withoutReviewGroupRecord(pkg, 'rg_retired');
  if (!removed.ok) throw new Error(removed.reason);
  expect(reviewGroupRevisionReservations(removed.pkg)).toEqual(
    reviewGroupRevisionReservations(pkg)
  );
  const stored = readReviewGroupStore(removed.pkg);
  expect(stored.status === 'ready' && stored.records.size).toBe(0);
  expect(stored.status === 'ready' && stored.retainedRecords.size).toBe(1);
  const editor = propertyEditor(removed.pkg);
  try {
    insertLeadingGroup(editor);
    expect(groups(editor)[0]!.native.every((member) => Number(member.address.id) > 100)).toBe(true);
  } finally {
    editor.destroy();
  }
});

test('ordinary native suggesting input and low-level session ops use the module reservation source', () => {
  const editor = propertyEditor(fixture());
  try {
    const surface = editor.surface!,
      paragraphId = groupParagraph(editor, 'The cat sat.');
    expect(editor.setEditingMode('suggesting').ok).toBe(true);
    surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 0 } });
    surface.insertPlainText('Typed ');
    surface.flushPendingInput();
    expect(
      collectRevisionSites(surface.session.part()).every(
        (site) =>
          Number(site.node.attributes.find((attribute) => attribute.localName === 'id')!.value) >
          100
      )
    ).toBe(true);
    const result = surface.session.applyTreeOps([
      {
        op: 'insertText',
        paragraphId,
        offset: 0,
        text: 'Native ',
        revision: { author: 'Other', date: '2026-10-20T10:00:00Z' },
      },
    ]);
    expect(result.rejected).toBe(false);
    expect(
      collectRevisionSites(surface.session.part()).every(
        (site) =>
          Number(site.node.attributes.find((attribute) => attribute.localName === 'id')!.value) >
          100
      )
    ).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(collectRevisionSites(surface.session.part())).toEqual([]);
  } finally {
    editor.destroy();
  }
});

for (const data of [
  '{"schemaVersion":3,"id":"rg_retired","future":{"keep":true}}',
  'malformed-json',
])
  test(`unsupported or malformed reservation records refuse before native writes: ${data}`, () => {
    const pkg = fixture(),
      store = readReviewGroupStore(pkg);
    if (store.status !== 'ready') throw new Error('Missing group store');
    const candidate = withCustomXmlNode(pkg, store.part.partName, {
        id: 'rg_retired',
        label: '',
        data,
      }),
      editor = propertyEditor(candidate);
    try {
      const session = editor.surface!.session,
        before = writeOoxmlPackage(session.currentPackage()),
        revision = session.packageRevision();
      expect(() => reviewGroupRevisionReservations(session.currentPackage())).toThrow();
      expect(() => proposeGroupText(editor, 'The cat sat.', 'The car sat.')).toThrow();
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
      expect(session.packageRevision()).toBe(revision);
      expect(session.canUndo()).toBe(false);
      expect(customXmlNodes(session.currentPackage(), store.part.partName)[0]!.data).toBe(data);
    } finally {
      editor.destroy();
    }
  });
