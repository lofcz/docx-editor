import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  readOoxmlPackage,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { propertyEditor } from '../../test/support/property-execution';
import { readReviewGroups } from './groups';
import { writeReviewGroupPackage } from './group-export';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { packageWithPropertyRecord, propertyRecord } from '../../test/support/property-record';
import { readPropertyReviewStore } from './property-record-store';
import { createSessionPackageWriter } from '../../../core/src/binding/session-package-writer';

test('save rebinds a pending cross-story group without changing live revision identities', async () => {
  const source = crossStoryGroupFixture();
  const fingerprints = () =>
    [...source.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]);
  const before = fingerprints();
  const editor = propertyEditor(source);
  try {
    const live = canonicalOoxmlFingerprint(
      editor.surface!.session.currentPackage().parts.get('/word/header1.xml')!
    );
    for (const bytes of [
      writeReviewGroupPackage(editor.surface!.session.currentPackage()),
      new Uint8Array(await editor.save()),
    ]) {
      const opened = readOoxmlPackage(bytes);
      if (!opened.ok) throw Error(opened.reason);
      const groups = readReviewGroups(opened.package);
      if (!groups.success) throw Error(groups.error);
      expect(
        groups.groups.find((group) => group.id === 'g1:stored:rg_cross_story')?.native
      ).toHaveLength(2);
      const ids = [...opened.package.parts.values()].flatMap((part) =>
        collectRevisionSites(part).map(
          (site) =>
            site.node.attributes.find(
              (a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'id'
            )!.value
        )
      );
      expect(new Set(ids).size).toBe(ids.length);
      const reopened = propertyEditor(opened.package);
      try {
        const twice = readOoxmlPackage(new Uint8Array(await reopened.save()));
        if (!twice.ok) throw Error(twice.reason);
        expect(
          [...twice.package.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)])
        ).toEqual(
          [...opened.package.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)])
        );
      } finally {
        reopened.destroy();
      }
    }
    expect(fingerprints()).toEqual(before);
    expect(
      canonicalOoxmlFingerprint(
        editor.surface!.session.currentPackage().parts.get('/word/header1.xml')!
      )
    ).toBe(live);
  } finally {
    editor.destroy();
  }
});

test('export rebinds property receipts and keeps unrelated extension data', () => {
  const payload = withPayloadRecord(
    crossStoryGroupFixture(),
    {
      version: 1,
      id: 'activity',
      type: 'lesson_activity',
      attributes: { duration: 10 },
    },
    'Activity'
  );
  if (!payload.ok) throw Error(payload.reason);
  const record = {
    ...propertyRecord(),
    nativeRevisions: [{ partName: '/word/header1.xml', id: '7', author: 'Teacher' }],
  };
  const source = packageWithPropertyRecord(payload.pkg, record);
  const read = readOoxmlPackage(writeReviewGroupPackage(source));
  if (!read.ok) throw Error(read.reason);
  const store = readPropertyReviewStore(read.package);
  if (store.status !== 'ready') throw Error('Missing exported property records');
  const decoded = store.records.get(record.id)!.decoded;
  if (decoded.status !== 'supported') throw Error('Unsupported exported property record');
  const native = collectRevisionSites(read.package.parts.get('/word/header1.xml')!)[0]!;
  const id = native.node.attributes.find(
    (a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'id'
  )!.value;
  expect(id).not.toBe('7');
  expect(decoded.value.nativeRevisions).toEqual([
    { partName: '/word/header1.xml', id, author: 'Teacher' },
  ]);
  expect(decoded.value.extension).toEqual(record.extension);
  const live = readPropertyReviewStore(source);
  if (live.status !== 'ready') throw Error('Missing live property records');
  expect(live.records.get(record.id)!.decoded).toEqual({ status: 'supported', value: record });
});

test('a refused export is not cached and successful cached bytes remain caller-owned', () => {
  const source = crossStoryGroupFixture();
  let refuse = true,
    calls = 0;
  const write = createSessionPackageWriter((_before, exported) => {
    calls += 1;
    if (refuse) throw Error('Unresolved module reference');
    return exported;
  });
  expect(() => write(source)).toThrow('Unresolved module reference');
  refuse = false;
  const first = write(source);
  const expected = first.slice();
  first.fill(0);
  const steps = write.steps(source);
  let next = steps.next();
  while (!next.done) next = steps.next();
  expect(next.value).toEqual(expected);
  expect(calls).toBe(2);
});
