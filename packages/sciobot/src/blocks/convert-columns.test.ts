import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  findNode,
  moveNode,
  paragraphTextOf,
  parentNodeOf,
  readOoxmlPackage,
  removeNode,
  semanticDigest,
  withPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { buildBlockIndex } from './block-index';
import { readColumnGeometry } from './column-geometry';
import { createColumnTableRefusal } from './column-table-owner';
import { convertColumnGroup, readLegacyColumnGroup } from './convert-columns';

async function legacyDocument(north = 'North', editXml?: (xml: string) => string) {
  const source = await targetDocumentInput('<w:p><w:r><w:t>Tail</w:t></w:r></w:p>');
  let id = 0;
  const planned = planTargetDocument(
    source,
    `<column_group>\n<column width="50%">\n\n${north}\n\n</column>\n<column width="50%">\n\nSouth\n\n</column>\n</column_group>\n\nTail`,
    { allocateBlockId: () => `legacy-${++id}` }
  );
  if (!planned.success) throw new Error(planned.error);
  const result = previewDirectPlan(source.pkg, planned.plan);
  if (!result.success) throw new Error(result.error);
  const group = [...buildBlockIndex(result.pkg).byId.values()].find(
    (block) => block.payload.type === 'column_group'
  )!;
  const geometry = readColumnGeometry(result.pkg, group)!;
  let part = result.pkg.parts.get(group.partName)!;
  const content = contentControlContentNodeOf(findNode(part, group.controlId)!)!;
  for (const [index, column] of geometry.columns.entries()) {
    const moved = moveNode(part, column.block.controlId, content.id, index);
    if (!moved.ok) throw new Error('Legacy column placement failed.');
    part = moved.part;
  }
  const removed = removeNode(part, geometry.tableId);
  if (!removed.ok) throw new Error('Legacy table removal failed.');
  part = removed.part;
  // The former stacked writer adds no terminal paragraph after a column control.
  for (const child of content.children.filter((node) => node.kind === 'paragraph')) {
    const removed = removeNode(part, child.id);
    if (!removed.ok) throw new Error('Legacy terminal paragraph removal failed.');
    part = removed.part;
  }
  let bytes = writeOoxmlPackage(withPart(result.pkg, part));
  if (editXml) {
    const files = unzipSync(bytes);
    files['word/document.xml'] = strToU8(editXml(strFromU8(files['word/document.xml']!)));
    bytes = zipSync(files);
  }
  return { bytes, groupId: group.id };
}
const request = (editor: DocxEditorInstance, groupId: string) => ({
  groupId,
  expectedRevision: editor.surface!.session.packageRevision(),
});
const geometry = (editor: DocxEditorInstance, groupId: string) => {
  const pkg = editor.surface!.session.currentPackage();
  return readColumnGeometry(pkg, buildBlockIndex(pkg).byId.get(groupId)!);
};

test('explicit legacy conversion retains rich native controls, text identities, metadata, save, and undo', async () => {
  const fixture = await legacyDocument(
    'North\n\n| A | B |\n| --- | --- |\n| One | **Two** |\n\n- First\n- Second'
  );
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fixture.bytes,
  });
  try {
    const surface = editor.surface!,
      pkg = surface.session.currentPackage();
    const legacy = readLegacyColumnGroup(pkg, fixture.groupId)!;
    expect(legacy.hasTrailingParagraph).toBe(false);
    const beforeProjection = (await projectSourcePackage(pkg)).text;
    const beforeDigest = semanticDigest(pkg.parts.values());
    const before = legacy.columns.map((column) =>
      canonicalOoxmlFingerprint(findNode(surface.session.part(), column.controlId)!)
    );
    const result = convertColumnGroup(editor, request(editor, fixture.groupId));
    if (!result.success) throw new Error(result.error);
    expect(geometry(editor, fixture.groupId)?.columns).toHaveLength(2);
    expect(
      legacy.columns.map((column) =>
        canonicalOoxmlFingerprint(findNode(surface.session.part(), column.controlId)!)
      )
    ).toEqual(before);
    expect(buildBlockIndex(surface.session.currentPackage()).byId).toEqual(
      buildBlockIndex(pkg).byId
    );
    expect((await projectSourcePackage(surface.session.currentPackage())).text).toBe(
      beforeProjection
    );
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(semanticDigest(saved.package.parts.values())).toEqual(
      semanticDigest(surface.session.currentPackage().parts.values())
    );
    expect((await projectSourcePackage(saved.package)).text).toBe(beforeProjection);
    expect(
      readColumnGeometry(saved.package, buildBlockIndex(saved.package).byId.get(fixture.groupId)!)
        ?.columns
    ).toHaveLength(2);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(semanticDigest(surface.session.currentPackage().parts.values())).toEqual(beforeDigest);
    expect(readLegacyColumnGroup(surface.session.currentPackage(), fixture.groupId)).not.toBeNull();
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(geometry(editor, fixture.groupId)).not.toBeNull();
  } finally {
    editor.destroy();
  }
});

test('conversion retains imported trailing paragraphs inside the last column', async () => {
  const fixture = await legacyDocument('North', (xml) =>
    xml.replace(/(<\/w:sdtContent><\/w:sdt>)(<w:p\b)/, '<w:p/>$1$2')
  );
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fixture.bytes,
  });
  try {
    const surface = editor.surface!;
    const legacy = readLegacyColumnGroup(surface.session.currentPackage(), fixture.groupId)!;
    expect(legacy.hasTrailingParagraph).toBe(true);
    const ids = surface.session.paragraphIds();
    const trailing = legacy.content.children.at(-1)!;
    const result = convertColumnGroup(editor, request(editor, fixture.groupId));
    if (!result.success) throw Error(result.error);
    expect(surface.session.paragraphIds()).toEqual(ids);
    const last = findNode(surface.session.part(), legacy.columns.at(-1)!.controlId)!;
    expect(parentNodeOf(surface.session.part(), trailing.id)?.id).toBe(
      contentControlContentNodeOf(last)!.id
    );
    expect(geometry(editor, fixture.groupId)?.separatorIds.size).toBe(0);
  } finally {
    editor.destroy();
  }
});

test('offline conversion repairs older native caret ownership without replacing identities', async () => {
  const fixture = await legacyDocument();
  const source = createDocxEditor({
    container: document.createElement('div'),
    document: fixture.bytes,
  });
  let bytes: Uint8Array;
  try {
    expect(convertColumnGroup(source, request(source, fixture.groupId)).success).toBe(true);
    const files = unzipSync(new Uint8Array(await source.save()));
    files['word/document.xml'] = strToU8(
      strFromU8(files['word/document.xml']!)
        .replaceAll('</w:sdt></w:tc>', '</w:sdt><w:p/></w:tc>')
        .replace('</w:tbl></w:sdtContent>', '</w:tbl><w:p/></w:sdtContent>')
    );
    bytes = zipSync(files);
  } finally {
    source.destroy();
  }
  const editor = createDocxEditor({ container: document.createElement('div'), document: bytes! });
  try {
    const surface = editor.surface!;
    const old = geometry(editor, fixture.groupId)!;
    expect(old.separatorIds.size).toBe(3);
    const ids = surface.session.paragraphIds();
    const before = semanticDigest(surface.session.currentPackage().parts.values());
    const index = buildBlockIndex(surface.session.currentPackage());
    const result = convertColumnGroup(editor, request(editor, fixture.groupId));
    if (!result.success) throw Error(result.error);
    expect(geometry(editor, fixture.groupId)?.separatorIds.size).toBe(0);
    expect(surface.session.paragraphIds()).toEqual(ids);
    expect(buildBlockIndex(surface.session.currentPackage()).byId).toEqual(index.byId);
    expect(geometry(editor, fixture.groupId)?.tableId).toBe(old.tableId);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(semanticDigest(surface.session.currentPackage().parts.values())).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw Error(saved.reason);
    expect(semanticDigest(saved.package.parts.values())).toEqual(
      semanticDigest(surface.session.currentPackage().parts.values())
    );
  } finally {
    editor.destroy();
  }
});

test('conversion refuses stale requests and non-editing modes without changing content', async () => {
  const fixture = await legacyDocument();
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fixture.bytes,
  });
  try {
    const surface = editor.surface!;
    const pkg = surface.session.currentPackage();
    expect(
      convertColumnGroup(editor, { ...request(editor, fixture.groupId), expectedRevision: -1 })
    ).toMatchObject({ success: false, code: 'stale-revision' });
    for (const mode of ['suggest', 'view'] as const) {
      surface.setEditingMode(mode);
      expect(convertColumnGroup(editor, request(editor, fixture.groupId))).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(surface.editingMode()).toBe(mode);
      expect(surface.session.currentPackage()).toBe(pkg);
    }
  } finally {
    editor.destroy();
  }
});

for (const [name, mutate] of [
  [
    'locked group',
    (xml: string) => xml.replace('<w:sdtPr>', '<w:sdtPr><w:lock w:val="contentLocked"/>'),
  ],
  [
    'bound columns',
    (xml: string) =>
      xml.replaceAll(
        '</w:sdtPr>',
        '<w:dataBinding w:storeItemID="{01}" w:xpath="/root/value"/></w:sdtPr>'
      ),
  ],
  [
    'locked columns',
    (xml: string) => xml.replaceAll('</w:sdtPr>', '<w:lock w:val="sdtLocked"/></w:sdtPr>'),
  ],
  [
    'extra content',
    (xml: string) =>
      xml.replace('<w:sdtContent>', '<w:sdtContent><w:p><w:r><w:t>Keep me</w:t></w:r></w:p>'),
  ],
] as const)
  test(`conversion refuses ${name} atomically`, async () => {
    const fixture = await legacyDocument('North', mutate);
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: fixture.bytes,
    });
    try {
      const pkg = editor.surface!.session.currentPackage();
      expect(convertColumnGroup(editor, request(editor, fixture.groupId))).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(editor.surface!.session.currentPackage()).toBe(pkg);
      expect(editor.surface!.session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

for (const first of [0, 1] as const)
  test(`shared legacy conversion refuses without changing either replica; receiver ${first}`, async () => {
    const fixture = await legacyDocument();
    const pair = await createTwoEditorFixture({ document: fixture.bytes });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const packages = pair.editors.map((editor) => editor.surface!.session.currentPackage());
      pair.pause();
      for (const [index, editor] of pair.editors.entries()) {
        expect(convertColumnGroup(editor, request(editor, fixture.groupId))).toMatchObject({
          success: false,
          code: 'unsupported-operation',
          error: 'Convert stacked columns before joining a collaboration session.',
        });
        expect(editor.surface!.session.currentPackage()).toBe(packages[index]);
      }
      const id = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'South')!;
      bob.surface!.setSelection({
        anchor: { paragraphId: id, offset: 5 },
        head: { paragraphId: id, offset: 5 },
      });
      bob.surface!.type(' Bob');
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(
          readLegacyColumnGroup(editor.surface!.session.currentPackage(), fixture.groupId)
        ).not.toBeNull();
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('South Bob');
      }
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(paragraphTextOf(alice.surface!.session.part(), id)).toBe('South');
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });

for (const first of [0, 1] as const)
  test(`offline converted columns join shared editing and preserve peer content through undo; receiver ${first}`, async () => {
    const fixture = await legacyDocument();
    const local = createDocxEditor({
      container: document.createElement('div'),
      document: fixture.bytes,
    });
    let documentBytes: Uint8Array;
    try {
      const result = convertColumnGroup(local, request(local, fixture.groupId));
      if (!result.success) throw new Error(result.error);
      documentBytes = new Uint8Array(await local.save());
    } finally {
      local.destroy();
    }
    const pair = await createTwoEditorFixture({
      document: documentBytes!,
      editor: { tableControlRefusal: createColumnTableRefusal('Use column controls.') },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const id = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'South')!;
      const north = alice
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'North')!;
      pair.pause();
      alice.surface!.setSelection({
        anchor: { paragraphId: north, offset: 5 },
        head: { paragraphId: north, offset: 5 },
      });
      alice.surface!.type(' Alice');
      bob.surface!.setSelection({
        anchor: { paragraphId: id, offset: 5 },
        head: { paragraphId: id, offset: 5 },
      });
      bob.surface!.type(' Bob');
      pair.reconnect(first);
      await pair.settle();
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(geometry(editor, fixture.groupId)?.columns).toHaveLength(2);
        expect(paragraphTextOf(editor.surface!.session.part(), north)).toBe('North');
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('South Bob');
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
