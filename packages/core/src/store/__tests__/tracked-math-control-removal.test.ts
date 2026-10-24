import { expect, test } from 'bun:test';
import {
  contentControlsIn,
  findNode,
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  applyTreeOp,
  collectRevisionSites,
  contentControlLevelOf,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
const ADA = { author: 'Ada Lovelace', date: '2026-01-02T03:04:05Z' };
function part(body: string): OoxmlPart {
  const read = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!read.ok) throw Error(read.reason);
  return read.part;
}
const xml = (source: OoxmlPart) => serializeOoxmlPart(source);
function apply(source: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(source, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
const paragraphId = (source: OoxmlPart) =>
  source.root.children[0]!.kind === 'textValue' ? '' : source.root.children[0]!.children[0]!.id;
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const equation = `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:sSup><m:e><m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`;
const control = (lock: string, mixed = true) =>
  `<w:sdt><w:sdtPr><w:tag w:val="app:math"/><w:lock w:val="${lock}"/><x:private xmlns:x="urn:test" x:keep="yes"/></w:sdtPr><w:sdtContent>${mixed ? run('[') : ''}${equation}${mixed ? run(']') : ''}</w:sdtContent></w:sdt>`;
for (const mixed of [false, true])
  for (const method of ['control', 'range'] as const)
    for (const linked of [false, true])
      for (const decision of ['accept', 'reject'] as const)
        test(`whole math control ${mixed}/${method}/${linked}/${decision} keeps exact equation metadata and decisions`, () => {
          const fragment = control(method === 'range' ? 'contentLocked' : 'unlocked', mixed);
          const source = part(
            `<w:p>${run('A ')}${linked ? `<w:hyperlink w:anchor="target">${fragment}</w:hyperlink>` : fragment}${run(' B')}</w:p>`
          );
          const original = contentControlsIn(source.root)[0]!.node;
          expect(contentControlLevelOf(original)).toBe('inline');
          const before = xml(source);
          const pending = apply(
            source,
            method === 'control'
              ? {
                  op: 'removeContentControl',
                  controlId: original.id,
                  keepContent: false,
                  revision: ADA,
                }
              : {
                  op: 'deleteText',
                  paragraphId: paragraphId(source),
                  start: 2,
                  end: mixed ? 5 : 3,
                  revision: ADA,
                }
          );
          const held = findNode(pending, original.id)!;
          const equationBefore = original.children
            .flatMap((child) => (child.kind === 'textValue' ? [] : child.children))
            .find((child) => child.kind !== 'textValue' && child.localName === 'oMath');
          expect(equationBefore).toBeDefined();
          expect(findNode(pending, equationBefore!.id)).toBe(equationBefore!);
          if (mixed)
            expect(canonicalOoxmlFingerprint(held)).not.toBe(canonicalOoxmlFingerprint(original));
          else expect(canonicalOoxmlFingerprint(held)).toBe(canonicalOoxmlFingerprint(original));
          const sites = collectRevisionSites(pending);
          expect(sites).toHaveLength(1);
          expect(sites[0]!.refused).toBe(false);
          for (const candidate of [
            pending,
            (() => {
              const read = readOoxmlPart(serializeOoxmlPart(pending), {
                name: pending.name,
                contentType: pending.contentType,
              });
              if (!read.ok) throw Error(read.reason);
              return read.part;
            })(),
          ]) {
            const decided = apply(candidate, {
              op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
              siteNodeIds: collectRevisionSites(candidate).map((site) => site.node.id),
            });
            const expected =
              decision === 'reject' ? source : part(`<w:p>${run('A ')}${run(' B')}</w:p>`);
            expect(canonicalOoxmlFingerprint(decided)).toBe(canonicalOoxmlFingerprint(expected));
          }
          expect(xml(source)).toBe(before);
        });

test('math control deletion still refuses nested native revision ownership without mutation', () => {
  const source = part(
    `<w:p>${control('unlocked').replace('<m:t>x</m:t>', '<w:ins w:id="2" w:author="Other"><m:t>x</m:t></w:ins>')}</w:p>`
  );
  const before = xml(source);
  expect(
    applyTreeOp(source, {
      op: 'removeContentControl',
      controlId: contentControlsIn(source.root)[0]!.node.id,
      keepContent: false,
      revision: ADA,
    })
  ).toMatchObject({ ok: false, reason: 'indivisible-content' });
  expect(xml(source)).toBe(before);
});
