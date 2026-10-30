/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { createRuntime } from '../runtime.ts';
import { openHost } from './support/hosts.ts';
import { docx, p } from './support/docx.ts';

const joined = (text: string) =>
  `<w:p><w:pPr><state:joined xmlns:state="urn:docx-editor.dev:review:retired-paragraph:1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const fixture = () => docx(joined('Left') + joined('Middle') + p('Tail') + p('Other'));

test('logical paragraph reads and ranges fold retained physical fragments without extra marks', async () => {
  const runtime = createRuntime({ host: openHost(fixture()), save: true });
  try {
    await runtime.run(async (c) => {
      const paragraphs = c.document.body.paragraphs;
      paragraphs.load('items/text');
      c.document.body.load('text');
      const range = paragraphs.getFirst().getRange('Whole');
      range.load('text');
      const covered = range.paragraphs;
      covered.load('items/text');
      await c.sync();
      expect(paragraphs.items.map((entry) => entry.text)).toEqual(['LeftMiddleTail', 'Other']);
      expect(c.document.body.text).toBe('LeftMiddleTail\rOther');
      expect(range.text).toBe('LeftMiddleTail');
      expect(covered.items.map((entry) => entry.text)).toEqual(['LeftMiddleTail']);
    });
  } finally {
    runtime.dispose();
  }
});

test('joined paragraph start and end writes land in the correct original physical fragments', async () => {
  const runtime = createRuntime({ host: openHost(fixture()), save: true });
  try {
    await runtime.run(async (c) => {
      c.document.body.paragraphs.getFirst().insertText('A', 'Start');
      await c.sync();
      c.document.body.paragraphs.getFirst().insertText('Z', 'End');
      await c.sync();
      const paragraphs = c.document.body.paragraphs;
      paragraphs.load('items/text');
      await c.sync();
      expect(paragraphs.items.map((entry) => entry.text)).toEqual(['ALeftMiddleTailZ', 'Other']);
    });
  } finally {
    runtime.dispose();
  }
});

test('joined paragraph search crosses retained marks and returns an editable range', async () => {
  const runtime = createRuntime({ host: openHost(fixture()), save: true });
  try {
    await runtime.run(async (c) => {
      const found = c.document.body.paragraphs.getFirst().search('ftMid');
      found.load('items/text');
      await c.sync();
      expect(found.items.map((entry) => entry.text)).toEqual(['ftMid']);
      found.items[0]!.insertText('X', 'Replace');
      await c.sync();
      const paragraphs = c.document.body.paragraphs;
      paragraphs.load('items/text');
      await c.sync();
      expect(paragraphs.items.map((entry) => entry.text)).toEqual(['LeXdleTail', 'Other']);
    });
  } finally {
    runtime.dispose();
  }
});

for (const where of ['Before', 'After'] as const)
  test(`inserting ${where} a joined paragraph uses its logical boundary`, async () => {
    const runtime = createRuntime({ host: openHost(fixture()), save: true });
    try {
      await runtime.run(async (c) => {
        c.document.body.paragraphs.getFirst().insertParagraph('Added', where);
        await c.sync();
        const paragraphs = c.document.body.paragraphs;
        paragraphs.load('items/text');
        await c.sync();
        expect(paragraphs.items.map((entry) => entry.text)).toEqual(
          where === 'Before'
            ? ['Added', 'LeftMiddleTail', 'Other']
            : ['LeftMiddleTail', 'Added', 'Other']
        );
      });
    } finally {
      runtime.dispose();
    }
  });

test('deleting a joined paragraph removes every fragment and preserves its next paragraph', async () => {
  const runtime = createRuntime({ host: openHost(fixture()), save: true });
  try {
    await runtime.run(async (c) => {
      c.document.body.paragraphs.getFirst().delete();
      await c.sync();
      const paragraphs = c.document.body.paragraphs;
      paragraphs.load('items/text');
      await c.sync();
      expect(paragraphs.items.map((entry) => entry.text)).toEqual(['Other']);
    });
  } finally {
    runtime.dispose();
  }
});

for (const trim of [false, true])
  test(`splitting a joined paragraph matches delimiters across fragments; trim=${trim}`, async () => {
    const runtime = createRuntime({ host: openHost(fixture()), save: true });
    try {
      await runtime.run(async (c) => {
        const pieces = c.document.body.paragraphs.getFirst().split(['ftMid'], trim);
        await c.sync();
        pieces.load('items/text');
        await c.sync();
        expect(pieces.items.map((entry) => entry.text)).toEqual(
          trim ? ['Le', 'dleTail'] : ['LeftMid', 'dleTail']
        );
        const paragraphs = c.document.body.paragraphs;
        paragraphs.load('items/text');
        await c.sync();
        expect(paragraphs.items.map((entry) => entry.text)).toEqual(
          trim ? ['Le', 'dleTail', 'Other'] : ['LeftMid', 'dleTail', 'Other']
        );
      });
    } finally {
      runtime.dispose();
    }
  });

test('replacing a complete joined paragraph keeps its logical proxy usable for later writes', async () => {
  const runtime = createRuntime({ host: openHost(fixture()), save: true });
  try {
    await runtime.run(async (c) => {
      const paragraph = c.document.body.paragraphs.getFirst();
      paragraph.insertText('Replacement', 'Replace');
      await c.sync();
      paragraph.insertText('!', 'End');
      await c.sync();
      paragraph.load('text');
      await c.sync();
      expect(paragraph.text).toBe('Replacement!');
    });
  } finally {
    runtime.dispose();
  }
});

test('a delimiter-free split trims its returned logical range without rewriting the document', async () => {
  // Word keeps boundary spaces in w:t only with xml:space="preserve".
  const preserved = (xml: string) => xml.replaceAll('<w:t>', '<w:t xml:space="preserve">');
  const runtime = createRuntime({
    host: openHost(docx(preserved(joined(' Left') + p('Tail ')))),
    save: true,
  });
  try {
    await runtime.run(async (c) => {
      const pieces = c.document.body.paragraphs.getFirst().split(['missing'], false, true);
      await c.sync();
      pieces.load('items/text');
      c.document.body.load('text');
      await c.sync();
      expect(pieces.items.map((entry) => entry.text)).toEqual(['LeftTail']);
      expect(c.document.body.text).toBe(' LeftTail ');
    });
  } finally {
    runtime.dispose();
  }
});

test('word ranges and reviewed text follow the logical paragraph across retained fragments', async () => {
  const runtime = createRuntime({
    host: openHost(docx(joined('one t') + p('wo three'))),
    save: true,
  });
  try {
    await runtime.run(async (c) => {
      const paragraph = c.document.body.paragraphs.getFirst();
      const words = paragraph.getTextRanges([' ']);
      words.load('items/text');
      const original = paragraph.getReviewedText('Original'),
        current = paragraph.getReviewedText('Current');
      await c.sync();
      expect(words.items.map((entry) => entry.text)).toEqual(['one ', 'two ', 'three']);
      expect(original.value).toBe('one two three');
      expect(current.value).toBe('one two three');
      words.items[1]!.insertText('new ', 'Replace');
      await c.sync();
      c.document.body.load('text');
      await c.sync();
      expect(c.document.body.text).toBe('one new three');
    });
  } finally {
    runtime.dispose();
  }
});
