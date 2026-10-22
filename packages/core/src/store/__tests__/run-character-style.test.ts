// `setRunProperties.characterStyle`: the explicit way a run write sets or removes `w:rStyle`.
//
// `w:rStyle` stays outside the accepted vocabulary, so a bold toggle keeps it. This field is
// the one channel that changes it, and everything else in each run's `w:rPr` must survive.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlPart } from '../package/ooxml-tree.ts';
import { applyTreeOp, validateTreeOp, type TreeDocOp } from '../store/tree-ops.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const X = 'urn:example:extension';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:x="${X}"><w:body>${body}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function paragraphId(part: OoxmlPart): string {
  const body = part.root.children.find((node) => node.kind !== 'textValue')!;
  if (body.kind === 'textValue') throw new Error('no body');
  return body.children.find((node) => node.kind === 'paragraph')!.id;
}

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return result.part;
}

function body(part: OoxmlPart): string {
  const xml = serializeOoxmlPart(part);
  return xml.slice(xml.indexOf('<w:body>') + 8, xml.indexOf('</w:body>'));
}

const SOURCE =
  '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:b/><w:lang w:val="cs-CZ"/><x:keep/></w:rPr>' +
  '<w:t>alpha beta</w:t></w:r></w:p>';

describe('setRunProperties.characterStyle', () => {
  test('sets a character style on a subrange and keeps every other run property', () => {
    const part = load(
      '<w:p><w:r><w:rPr><w:b/><w:lang w:val="cs-CZ"/><x:keep/></w:rPr><w:t>alpha beta</w:t></w:r></w:p>'
    );
    const id = paragraphId(part);
    const next = apply(part, {
      op: 'setRunProperties',
      paragraphId: id,
      start: 6,
      end: 10,
      properties: [{ localName: 'b' }],
      characterStyle: 'SciobotCode',
    });
    expect(body(next)).toBe(
      '<w:p><w:r><w:rPr><w:b/><w:lang w:val="cs-CZ"/><x:keep/></w:rPr><w:t xml:space="preserve">alpha </w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="SciobotCode"/><w:b/><w:lang w:val="cs-CZ"/><x:keep/></w:rPr><w:t>beta</w:t></w:r></w:p>'
    );
  });

  test('replaces and removes an existing reference without touching other properties', () => {
    const part = load(SOURCE);
    const id = paragraphId(part);
    const replaced = apply(part, {
      op: 'setRunProperties',
      paragraphId: id,
      start: 0,
      end: 10,
      properties: [{ localName: 'b' }],
      characterStyle: 'SciobotCode',
    });
    expect(body(replaced)).toContain(
      '<w:rPr><w:rStyle w:val="SciobotCode"/><w:b/><w:lang w:val="cs-CZ"/><x:keep/></w:rPr>'
    );
    const removed = apply(replaced, {
      op: 'setRunProperties',
      paragraphId: id,
      start: 0,
      end: 10,
      properties: [{ localName: 'b' }],
      characterStyle: null,
    });
    expect(body(removed)).toBe(
      '<w:p><w:r><w:rPr><w:b/><w:lang w:val="cs-CZ"/><x:keep/></w:rPr><w:t>alpha beta</w:t></w:r></w:p>'
    );
  });

  test('omitting the field keeps the authored reference, as before', () => {
    const part = load(SOURCE);
    const next = apply(part, {
      op: 'setRunProperties',
      paragraphId: paragraphId(part),
      start: 0,
      end: 10,
      properties: [],
    });
    expect(body(next)).toBe(
      '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:lang w:val="cs-CZ"/><x:keep/></w:rPr><w:t>alpha beta</w:t></w:r></w:p>'
    );
  });

  test('an equal reference keeps its node identity', () => {
    const part = load(SOURCE);
    const id = paragraphId(part);
    const before = serializeOoxmlPart(part);
    const next = apply(part, {
      op: 'setRunProperties',
      paragraphId: id,
      start: 0,
      end: 10,
      properties: [{ localName: 'b' }],
      characterStyle: 'Emphasis',
    });
    expect(serializeOoxmlPart(next)).toBe(before);
  });

  test('a tracked change records the prior reference so rejection restores it', () => {
    const part = load(SOURCE);
    const id = paragraphId(part);
    const revision = { author: 'ScioBot', date: '2026-10-21T08:00:00Z' };
    const tracked = apply(part, {
      op: 'setRunProperties',
      paragraphId: id,
      start: 0,
      end: 5,
      properties: [{ localName: 'b' }],
      characterStyle: 'SciobotCode',
      revision,
    });
    const xml = body(tracked);
    expect(xml).toContain('<w:rStyle w:val="SciobotCode"/>');
    expect(xml).toMatch(
      /<w:rPrChange w:author="ScioBot" w:date="2026-10-21T08:00:00Z" w:id="(\d+)"><w:rPr><w:rStyle w:val="Emphasis"\/><w:b\/><w:lang w:val="cs-CZ"\/><x:keep\/><\/w:rPr><\/w:rPrChange>/
    );
    const revisionId = /<w:rPrChange [^>]*w:id="(\d+)"/.exec(xml)![1]!;
    const rejected = apply(tracked, {
      op: 'rejectRevision',
      revision: { id: revisionId, author: 'ScioBot', date: '2026-10-21T08:00:00Z' },
    } as TreeDocOp);
    expect(body(rejected)).not.toContain('SciobotCode');
    expect(body(rejected)).not.toContain('rPrChange');
    const accepted = apply(tracked, {
      op: 'acceptRevision',
      revision: { id: revisionId, author: 'ScioBot', date: '2026-10-21T08:00:00Z' },
    } as TreeDocOp);
    expect(body(accepted)).toContain('<w:rStyle w:val="SciobotCode"/>');
    expect(body(accepted)).not.toContain('rPrChange');
  });

  test('malformed style IDs refuse before any write', () => {
    const part = load(SOURCE);
    const id = paragraphId(part);
    for (const characterStyle of ['', 'a\u0000b', 'x'.repeat(513), 7, {}] as unknown[]) {
      const op = {
        op: 'setRunProperties',
        paragraphId: id,
        start: 0,
        end: 5,
        properties: [],
        characterStyle,
      } as unknown as TreeDocOp;
      expect(validateTreeOp(part, op)).toBe('invalid-property-value');
      expect(applyTreeOp(part, op).ok).toBe(false);
    }
  });
});
