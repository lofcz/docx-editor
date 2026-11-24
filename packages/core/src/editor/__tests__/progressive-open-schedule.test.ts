// The prefixes a large open lays out are sized by content, so every machine lays out the same
// prefixes and publishes the same pagination.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/package/ooxml-tree.ts';
import {
  bodyBlockWeight,
  nextPrefixLimit,
  preparedParagraphPrefix,
} from '../surface-progressive-open.ts';

const text = (value: string) => ({ kind: 'textValue', value });
const paragraph = (length: number) => ({ kind: 'paragraph', children: [text('x'.repeat(length))] });

test('prepared paragraph openings retain the original policy for short and mixed bodies', () => {
  function prefix(body: string): number | null {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`;
    const loaded = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'app/xml' });
    if (!loaded.ok) throw new Error(loaded.reason);
    return preparedParagraphPrefix(loaded.part.root);
  }
  expect(prefix('<w:p/>'.repeat(32))).toBeNull();
  expect(prefix('<w:p/>'.repeat(64))).toBeNull();
  expect(prefix('<w:p/>'.repeat(65))).toBe(16);
  expect(prefix('<w:p/>'.repeat(320))).toBe(16);
  expect(prefix('<w:p/>'.repeat(320) + '<w:tbl/>')).toBeNull();
  expect(prefix('<w:p/>'.repeat(70) + '<w:sectPr/><w:p/>')).toBeNull();
  expect(prefix('<w:p/>'.repeat(70) + '<other:sectPr xmlns:other="urn:example"/>')).toBeNull();
});

describe('nextPrefixLimit', () => {
  test('a step holds a fixed amount of text, at least a few blocks', () => {
    const blocks = Array.from({ length: 10_000 }, () => paragraph(1_000));
    const first = nextPrefixLimit(0, blocks.length, (index) => blocks[index]);
    expect(first).toBe(60);
    expect(nextPrefixLimit(first, blocks.length, (index) => blocks[index])).toBe(120);
    const short = Array.from({ length: 100 }, () => paragraph(1));
    expect(nextPrefixLimit(0, short.length, (index) => short[index])).toBe(100);
  });

  test('heavy blocks still advance by the minimum step, light ones by the maximum', () => {
    const heavy = Array.from({ length: 100 }, () => paragraph(500_000));
    expect(nextPrefixLimit(10, heavy.length, (index) => heavy[index])).toBe(26);
    const light = Array.from({ length: 20_000 }, () => paragraph(0));
    expect(nextPrefixLimit(0, light.length, (index) => light[index])).toBe(4096);
  });

  test('the weight of a block is its text, counted once per block', () => {
    const block = { kind: 'table', children: [paragraph(10), paragraph(20)] };
    expect(bodyBlockWeight(block)).toBe(31);
    expect(bodyBlockWeight(block)).toBe(31);
  });
});
