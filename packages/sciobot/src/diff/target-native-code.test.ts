import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { codeDocument, openCode } from '../../test/support/code';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import { buildBlockIndex } from '../blocks/block-index';
import { parseTargetMarkdown } from '../projection/parse-target';
import { atomicTargetsPairable } from './target-atomic';
import { planTargetDocument } from './plan-target';
import { planTargetNativeCode } from './target-native-code';

test('prepared code target routing refuses stale, foreign, unsupported, and version-1 source owners', async () => {
  const pkg = openCode(await codeDocument());
  const projection = await projectSourcePackage(pkg);
  const payload = buildBlockIndex(pkg).byId.get('code')!.payload;
  const parsed = parseTargetMarkdown(projection.text);
  if (!parsed.success) throw Error(parsed.error);
  const before = parsed.target.ast.children[0]!;
  const after = { ...before, value: 'changed' };
  const source = { pkg, revision: sourceRevision, blockId: 'code' };
  const snapshot = writeOoxmlPackage(pkg);
  for (const [owner, old, target] of [
    [{ ...payload, id: 'foreign' }, before, after],
    [payload, { ...before, value: 'stale' }, after],
    [payload, before, { ...after, meta: 'unsupported' }],
    [payload, before, { ...after, lang: ['not a language'] }],
    [{ ...payload, version: 1 }, before, after],
  ] as const) {
    expect(planTargetNativeCode(source, owner, old, target)).toMatchObject({ success: false });
    expect(writeOoxmlPackage(pkg)).toEqual(snapshot);
  }
  expect(atomicTargetsPairable({ ...payload, version: 1 }, before, after)).toBe(false);
  const planned = planTargetDocument(
    { pkg, revision: sourceRevision, projection },
    projection.text.replace('middle', 'changed')
  );
  if (!planned.success) throw Error(planned.error);
  expect(planned.plan.operations).toHaveLength(1);
  expect(planned.plan.operations[0]).toMatchObject({ kind: 'payload', blockId: 'code', path: [] });
});

test('complete code targets retain imported locks and bindings, including mixed prose edits', async () => {
  const bytes = await codeDocument();
  for (const protection of ['unowned', 'sdtLocked', 'sdtContentLocked', 'ancestor', 'binding']) {
    const entries = unzipSync(bytes);
    let xml = strFromU8(entries['word/document.xml']!);
    if (protection === 'unowned')
      xml = xml.replace('scioCode:sourceOwner="1"', 'scioCode:sourceOwner="0"');
    else if (protection === 'sdtLocked' || protection === 'sdtContentLocked')
      xml = xml.replace('w:val="contentLocked"', `w:val="${protection}"`);
    else {
      const properties =
        protection === 'ancestor'
          ? '<w:lock w:val="contentLocked"/>'
          : '<w:dataBinding w:xpath="/value" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/>';
      xml = xml
        .replace('<w:body>', `<w:body><w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>`)
        .replace('</w:body>', '</w:sdtContent></w:sdt></w:body>');
    }
    entries['word/document.xml'] = strToU8(xml);
    const pkg = openCode(zipSync(entries));
    const projection = await projectSourcePackage(pkg);
    const snapshot = writeOoxmlPackage(pkg);
    expect(
      planTargetDocument(
        { pkg, revision: sourceRevision, projection },
        projection.text.replace('middle', 'changed').replace('Tail', 'New tail')
      )
    ).toMatchObject({ success: false });
    expect(writeOoxmlPackage(pkg)).toEqual(snapshot);
  }
});
