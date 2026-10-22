import { expect, test } from 'bun:test';
import {
  readOoxmlPart,
  serializeOoxmlPart,
  withPart,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { partReviewFieldBudget, groupReviewFieldBudget } from './group-field-budget';

function part(attributes = '') {
  const read = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test">' +
      '<w:body><w:p><w:r><w:t ' +
      attributes +
      '>Keep</w:t></w:r></w:p></w:body></w:document>',
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  return read.part;
}

test('field headroom includes removed subtree names, attributes, and namespace declarations', () => {
  const before = part(Array.from({ length: 100 }, (_, i) => 'x:a' + i + '="v"').join(' ')),
    bytes = serializeOoxmlPart(before),
    empty: OoxmlPart = { ...before, root: { ...before.root, children: [] as const } };
  expect(partReviewFieldBudget(before, before)).toBe(0);
  expect(partReviewFieldBudget(before, empty)).toBeGreaterThanOrEqual(200);
  expect(partReviewFieldBudget(before, undefined)).toBeGreaterThanOrEqual(200);
  expect(partReviewFieldBudget(undefined, before)).toBeGreaterThanOrEqual(200);
  expect(serializeOoxmlPart(before)).toBe(bytes);
});

test('field headroom covers renamed descriptors and transient writes even when the final package is unchanged', () => {
  const source = crossStoryGroupFixture(),
    before = source.parts.get(source.mainDocumentPart)!,
    changed = {
      ...before,
      root: {
        ...before.root,
        prefix: 'word',
        namespaceBindings: [
          ...before.root.namespaceBindings,
          { prefix: 'word', namespaceUri: before.root.namespaceUri },
        ],
      },
    },
    after = withPart(source, changed);
  expect(partReviewFieldBudget(before, changed)).toBeGreaterThanOrEqual(18);
  expect(groupReviewFieldBudget(source, source)).toBe(0);
  expect(
    groupReviewFieldBudget(source, after) + groupReviewFieldBudget(after, source)
  ).toBeGreaterThan(0);
});

test('field headroom includes package relationships and content-type records outside document parts', () => {
  const source = crossStoryGroupFixture(),
    after = {
      ...source,
      relationships: new Map(source.relationships),
      contentTypes: {
        defaults: new Map(source.contentTypes.defaults),
        overrides: new Map(source.contentTypes.overrides),
      },
    };
  after.relationships.set('/word/document.xml', [
    {
      id: 'rId999',
      type: 'urn:test',
      ownerPart: '/word/document.xml',
      rawTarget: 'media/test.png',
      targetMode: 'Internal',
      order: 0,
    },
  ]);
  after.contentTypes.defaults.set('probe', 'application/probe');
  expect(groupReviewFieldBudget(source, after)).toBeGreaterThanOrEqual(16);
});
