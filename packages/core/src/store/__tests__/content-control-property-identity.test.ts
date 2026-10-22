// A metadata write keeps the property element it changes.
//
// Collaboration replicates canonical identities. A fresh `w:tag` per write lets two concurrent
// writers each remove the old element and insert their own, so the merged control holds two
// tags and no reader accepts it. Changing the value in place makes the write one attribute.

import { expect, test } from 'bun:test';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parse(sdtPr: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:sdt><w:sdtPr>${sdtPr}</w:sdtPr>` +
      `<w:sdtContent><w:p><w:r><w:t>v</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function control(part: OoxmlPart): Exclude<OoxmlNode, { kind: 'textValue' }> {
  return contentControlsIn(part.root)[0]!.node as Exclude<OoxmlNode, { kind: 'textValue' }>;
}

function propertyIds(part: OoxmlPart): Record<string, string> {
  const sdtPr = control(part).children[0] as Exclude<OoxmlNode, { kind: 'textValue' }>;
  return Object.fromEntries(
    sdtPr.children.flatMap((child) =>
      child.kind === 'textValue' ? [] : [[child.localName, child.id]]
    )
  );
}

function apply(
  part: OoxmlPart,
  op: Omit<TreeDocOp & { op: 'setContentControlProperties' }, 'op' | 'controlId'>
) {
  const result = applyTreeOp(part, {
    op: 'setContentControlProperties',
    controlId: control(part).id,
    ...op,
  } as TreeDocOp);
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

test('changing a tag or alias keeps the property element and its other attributes', () => {
  const part = parse(
    '<w:alias w:val="Old" w14:extra="kept" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>' +
      '<w:tag w:val="old-tag"/><w:id w:val="7"/>'
  );
  const before = propertyIds(part);
  const next = apply(part, { tag: 'new-tag', alias: 'New' });
  expect(propertyIds(next)).toEqual(before);
  const properties = contentControlPropertiesOf(control(next));
  expect([properties.tag, properties.alias, properties.id]).toEqual(['new-tag', 'New', 7]);
  expect(serializeOoxmlPart(next)).toContain('w14:extra="kept"');
});

test('an unchanged value leaves the property object untouched', () => {
  const part = parse('<w:tag w:val="same"/>');
  const sdtPr = control(part).children[0] as Exclude<OoxmlNode, { kind: 'textValue' }>;
  const next = apply(part, { tag: 'same' });
  const after = control(next).children[0] as Exclude<OoxmlNode, { kind: 'textValue' }>;
  expect(after.children[0]).toBe(sdtPr.children[0]);
});

test('a missing property is still added, and removal still drops it', () => {
  const part = parse('<w:tag w:val="t"/>');
  const added = apply(part, { alias: 'Shown', lock: 'sdtLocked' });
  expect(contentControlPropertiesOf(control(added))).toMatchObject({
    tag: 't',
    alias: 'Shown',
    lock: 'sdtLocked',
  });
  const unlocked = apply(added, { lock: 'unlocked' });
  expect(contentControlPropertiesOf(control(unlocked)).lock).toBe('unlocked');
  expect(propertyIds(unlocked).lock).toBeUndefined();
  expect(propertyIds(unlocked).alias).toBe(propertyIds(added).alias);
});
