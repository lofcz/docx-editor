import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../index.ts';
import { ADA, apply, paragraphId, part } from './tracked-edit-fixture.ts';

const history = '<w:rPrChange w:id="42" w:author="Other"><w:rPr><w:i/></w:rPr></w:rPrChange>';
const fixtures = [
  { name: 'absent', source: '', expected: '' },
  { name: 'literal empty', source: '<w:rPr/>', expected: '<w:rPr/>' },
  {
    name: 'empty with extension attributes',
    source: '<w:rPr xmlns:keep="urn:keep" keep:face="literal"/>',
    expected: '<w:rPr xmlns:keep="urn:keep" keep:face="literal"/>',
  },
  { name: 'history only', source: `<w:rPr>${history}</w:rPr>`, expected: '' },
  {
    name: 'face with foreign history',
    source: `<w:rPr><w:b/>${history}</w:rPr>`,
    expected: '<w:rPr><w:b/></w:rPr>',
  },
];
const children = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const properties = (run: OoxmlNode) =>
  children(run)
    .filter((node) => node.kind === 'runProperties')
    .map(canonicalOoxmlFingerprint);
function reopen(source: OoxmlPart): OoxmlPart {
  const opened = readOoxmlPart(serializeOoxmlPart(source), {
    name: source.name,
    contentType: source.contentType,
  });
  if (!opened.ok) throw Error(opened.reason);
  return opened.part;
}

for (const fixture of fixtures)
  for (const offset of [0, 2, 4])
    for (const operation of ['insertText', 'insertTab', 'insertHardBreak'] as const)
      test(`${operation} at ${offset} inherits ${fixture.name} run properties without foreign history`, () => {
        const source = part(`<w:p><w:r>${fixture.source}<w:t>word</w:t></w:r></w:p>`);
        const before = canonicalOoxmlFingerprint(source);
        const expected = part(`<w:p><w:r>${fixture.expected}<w:t>X</w:t></w:r></w:p>`);
        const expectedRun = children(findNode(expected, paragraphId(expected))!)[0]!;
        const position = { paragraphId: paragraphId(source), offset, revision: ADA };
        const pending = apply(
          source,
          operation === 'insertText'
            ? { op: operation, ...position, text: 'X' }
            : { op: operation, ...position }
        );
        for (const current of [pending, reopen(pending)]) {
          const insertion = collectRevisionSites(current).find(
            (site) => site.node.kind === 'revisionInsert'
          );
          if (!insertion) throw Error('Missing native insertion.');
          const runs = children(insertion.node).filter((node) => node.kind === 'run');
          expect(runs).toHaveLength(1);
          expect(properties(runs[0]!)).toEqual(properties(expectedRun));
          const attribute = (name: string) =>
            insertion.node.attributes.find((item) => item.localName === name)?.value;
          const accepted = apply(current, {
            op: 'acceptRevision',
            revision: { id: attribute('id')!, author: ADA.author, date: ADA.date },
            siteNodeIds: [insertion.node.id],
          });
          expect(properties(findNode(accepted, runs[0]!.id)!)).toEqual(properties(expectedRun));
          expect(
            collectRevisionSites(accepted).some((site) => site.node.kind === 'revisionInsert')
          ).toBe(false);
        }
        expect(canonicalOoxmlFingerprint(source)).toBe(before);
      });
