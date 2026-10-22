import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  storyParagraphs,
  bodyStoryRoot,
} from '@docx-editor.dev/core/store';
import { bodyPart } from '../../test/support/document';
import { nativePointAt, nativeRangeAt, projectBodyText } from './text';

const revision = { documentId: 'lesson-1', epoch: 'mount-1', revision: 3 };
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
function project(body: string) {
  const part = bodyPart(body);
  const result = projectBodyText(part, revision);
  if (!result.success) throw new Error(result.error);
  return { part, projection: result.projection };
}

test('Czech accents and emoji retain exact native UTF-16 ranges', () => {
  const { projection } = project(`<w:p>${run('Žák 😀 čeká.')}</w:p>`);
  expect(projection.text).toBe('Žák 😀 čeká.');
  const id = projection.paragraphs[0]!.id;
  expect(nativeRangeAt(projection, 4, 6)).toEqual({
    start: { paragraphId: id, offset: 4 },
    end: { paragraphId: id, offset: 6 },
  });
  expect(nativeRangeAt(projection, 5, 6)).toBeNull();
  expect(nativePointAt(projection, 5)).toBeNull();
});

test('tabs and native line breaks stay within the same paragraph offset space', () => {
  const { projection } = project(
    `<w:p>${run('A')}<w:r><w:tab/></w:r>${run('B')}<w:r><w:br/></w:r>${run('C')}</w:p>`
  );
  expect(projection.text).toBe('A\tB\nC');
  const id = projection.paragraphs[0]!.id;
  expect(nativeRangeAt(projection, 1, 4)).toEqual({
    start: { paragraphId: id, offset: 1 },
    end: { paragraphId: id, offset: 4 },
  });
});

test('carriage returns and page breaks preserve their native UTF-16 units', () => {
  const { projection } = project(
    `<w:p>${run('A')}<w:r><w:cr/></w:r>${run('B')}<w:r><w:br w:type="page"/></w:r>${run('C')}</w:p>`
  );
  expect(projection.text).toBe('A\nB\fC');
  expect(nativeRangeAt(projection, 2, 5)).toEqual({
    start: { paragraphId: projection.paragraphs[0]!.id, offset: 2 },
    end: { paragraphId: projection.paragraphs[0]!.id, offset: 5 },
  });
});

test('empty paragraphs keep identities and generated separators reject native targeting', () => {
  const { projection } = project(`<w:p>${run('A')}</w:p><w:p/><w:p>${run('B')}</w:p>`);
  expect(projection.text).toBe('A\n\n\n\nB');
  expect(projection.paragraphs).toHaveLength(3);
  expect(nativePointAt(projection, 3)).toEqual({
    paragraphId: projection.paragraphs[1]!.id,
    offset: 0,
  });
  expect(nativePointAt(projection, 2)).toBeNull();
  expect(nativeRangeAt(projection, -1, 1)).toBeNull();
  expect(nativeRangeAt(projection, 3, 1)).toBeNull();
  expect(nativePointAt(projection, 1.5)).toBeNull();
  expect(nativePointAt(projection, Infinity)).toBeNull();
});

test('repeated text resolves to distinct paragraph identities', () => {
  const { projection } = project(`<w:p>${run('Same')}</w:p><w:p>${run('Same')}</w:p>`);
  expect(nativeRangeAt(projection, 6, 10)?.start).toEqual({
    paragraphId: projection.paragraphs[1]!.id,
    offset: 0,
  });
  expect(projection.paragraphs[0]!.id).not.toBe(projection.paragraphs[1]!.id);
});

test('native containers and tables retain body reading order without changing the part', () => {
  const { part, projection } = project(
    `<w:sdt><w:sdtPr/><w:sdtContent><w:p>${run('Inside')}</w:p></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc><w:p>${run('Cell')}</w:p></w:tc></w:tr></w:tbl>`
  );
  const original = canonicalOoxmlFingerprint(part);
  expect(projection.text).toBe('Inside\n\nCell');
  expect(projection.paragraphs.map((paragraph) => paragraph.id)).toEqual(
    storyParagraphs(bodyStoryRoot(part)!).map((paragraph) => paragraph.id)
  );
  const token = { ...revision };
  const again = projectBodyText(part, token);
  token.revision = 9;
  expect(again.success && again.projection.revision.revision).toBe(3);
  expect(canonicalOoxmlFingerprint(part)).toBe(original);
  expect(JSON.parse(JSON.stringify(projection))).toEqual(projection);
});
