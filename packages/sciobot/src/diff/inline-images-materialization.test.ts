import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  resolveImageRelationship,
  type OoxmlPackage,
  type ImageDecodePort,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from './plan-target';
const context = { documentId: 'inline-images', epoch: 'native' };
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => /(^|\/)media\//.test(name));
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  let next = 0;
  const result = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target,
    { allocateBlockId: () => `control-${++next}` }
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
for (const target of [
  'Keep.\n\nA ![Cat](asset:cat "Photo") Z',
  'Keep. A ![Cat](asset:cat "Photo") Z',
  'Keep. [Link](https://example.test)![Cat](asset:cat "Photo") Z',
  'Keep.\n\n<phase type="evocation">\nA ![Cat](asset:cat "Photo") Z\n</phase>',
  'Keep.\n\n| H |\n| - |\n| A ![Cat](asset:cat "Photo") Z |',
  'Keep. A $x$ ![Cat](asset:cat "Photo") [Ada](mention:user) Z',
  'Keep.\n\nA ![Cat](asset:cat "Photo") ![Dog](asset:dog) Z',
  'Keep. A ![' + 'Long alt '.repeat(600) + '](asset:cat) Z',
])
  test(`new inline drawings retain source references, media, native history and DOCX: ${target.slice(0, 130)}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      imageDecodePort: decode,
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const originalId = editor.surface!.session.paragraphIds()[0]!;
      const planned = await plan(editor, target);
      const original = editor.surface!.session.currentPackage();
      const resolved: string[] = [];
      const prepared = await prepareInlineImageResources(
        planned,
        async (reference) => {
          resolved.push(reference);
          return { bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 };
        },
        decode
      );
      if (!prepared.success) throw Error(prepared.error);
      expect(editor.surface!.session.currentPackage()).toBe(original);
      expect(media(original)).toEqual([]);
      expect(resolved).toEqual(
        target.includes('asset:dog') ? ['asset:cat', 'asset:dog'] : ['asset:cat']
      );
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const applied = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
      stop();
      expect(publications).toBe(1);
      if (!applied.success) throw Error(applied.error);
      expect(editor.surface!.session.paragraphIds()).toContain(originalId);
      const pkg = editor.surface!.session.currentPackage();
      const index = buildBlockIndex(pkg);
      expect(index.diagnostics).toEqual([]);
      const images = [...index.byId.values()].filter(
        (block) => block.payload.type === 'inline_image'
      );
      expect(images).toHaveLength(target.includes('asset:dog') ? 2 : 1);
      expect(media(pkg)).toHaveLength(images.length);
      for (const [, bytes] of media(pkg)) expect(bytes).toEqual(png);
      for (const image of images) {
        expect(image.payload.attributes.widthPoints).toBe(12);
        expect(image.payload.attributes.heightPoints).toBe(8);
        expect(String(image.payload.attributes.contentHash)).toHaveLength(64);
      }
      expect((await plan(editor, target)).operations).toEqual([]);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(media(editor.surface!.session.currentPackage())).toEqual([]);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      expect(media(editor.surface!.session.currentPackage())).toEqual(media(pkg));
      const bytes = new Uint8Array(await editor.save());
      const loaded = readOoxmlPackage(bytes);
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      expect(media(loaded.package)).toEqual(media(pkg));
      expect(
        [...buildBlockIndex(loaded.package).byId].map(([id, block]) => [id, block.payload])
      ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
        imageDecodePort: decode,
      });
      try {
        expect((await plan(reopened, target)).operations).toEqual([]);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

test('later edits keep native image identities and bytes; only newly inserted images resolve and allocate media', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  const first = 'Keep. A ![Cat](asset:cat "Photo") Z';
  const second = 'Prefix Keep. A ![Cat](asset:cat "Photo") changed ![Dog](asset:dog) Z';
  const third = second.replace('Prefix', 'Updated');
  const insert = async (target: string) => {
    const planned = await plan(editor, target),
      resolved: string[] = [];
    const prepared = await prepareInlineImageResources(
      planned,
      async (reference) => {
        resolved.push(reference);
        return { bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 };
      },
      decode
    );
    if (!prepared.success) throw Error(prepared.error);
    const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
    if (!result.success) throw Error(result.error);
    return resolved;
  };
  try {
    expect(await insert(first)).toEqual(['asset:cat']);
    const pkg = editor.surface!.session.currentPackage();
    const cat = [...buildBlockIndex(pkg).byId.values()][0]!;
    const catWord = canonicalOoxmlFingerprint(
      findNode(editor.surface!.session.part(), cat.controlId)!
    );
    const oldMedia = media(pkg);
    expect(await insert(second)).toEqual(['asset:dog']);
    const afterSecond = canonicalOoxmlFingerprint(editor.surface!.session.part());
    expect(media(editor.surface!.session.currentPackage())).toHaveLength(2);
    for (const [name, bytes] of oldMedia)
      expect(editor.surface!.session.currentPackage().partBytes.get(name)).toEqual(bytes);
    expect(await insert(third)).toEqual([]);
    const current = editor.surface!.session.currentPackage();
    const currentCat = buildBlockIndex(current).byId.get(cat.id)!;
    expect(currentCat.controlId).toBe(cat.controlId);
    expect(currentCat.payload).toEqual(cat.payload);
    expect(
      canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), cat.controlId)!)
    ).toBe(catWord);
    expect((await plan(editor, third)).operations).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(afterSecond);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(media(editor.surface!.session.currentPackage())).toEqual(oldMedia);
    expect(buildBlockIndex(editor.surface!.session.currentPackage()).diagnostics).toEqual([]);
    expect(
      buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(cat.id)?.payload
    ).toEqual(cat.payload);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw Error(reopened.reason);
    expect(
      canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
    ).toBe(canonicalOoxmlFingerprint(editor.surface!.session.part()));
    expect(media(reopened.package)).toEqual(media(current));
  } finally {
    editor.destroy();
  }
});

test('prepared images snapshot bytes and metadata, reject forged or wrong-plan handles, and retain atomic publication', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    const planned = await plan(editor, 'Keep. A ![Cat](asset:cat) Z');
    const other = await plan(editor, 'Keep. A ![Dog](asset:dog) Z');
    const source = {
      bytes: new Uint8Array(png),
      mime: 'image/png' as const,
      widthPoints: 12,
      heightPoints: 8,
    };
    const prepared = await prepareInlineImageResources(planned, async () => source, {
      async decode(copy) {
        copy.fill(0);
        source.bytes.fill(0);
        source.widthPoints = 999;
        return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
      },
    });
    if (!prepared.success) throw Error(prepared.error);
    const before = editor.surface!.session.currentPackage();
    for (const [input, handle] of [
      [planned, { kind: 'prepared-inline-images' as const }],
      [other, prepared.images],
    ] as const) {
      expect(applyDirectPlan(editor, input, { ...context, images: handle })).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
    }
    const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
    if (!result.success) throw Error(result.error);
    const pkg = editor.surface!.session.currentPackage();
    expect(media(pkg)[0]![1]).toEqual(png);
    expect([...buildBlockIndex(pkg).byId.values()][0]!.payload.attributes.widthPoints).toBe(12);
  } finally {
    editor.destroy();
  }
});

test('image resource failures and cancellation preserve the complete package and history', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    const planned = await plan(editor, 'Keep. A ![Cat](asset:cat) Z');
    const before = editor.surface!.session.currentPackage();
    for (const kind of ['invalid-bytes', 'mime', 'decode', 'dimensions', 'resolver', 'abort']) {
      const abort = new AbortController();
      if (kind === 'abort') abort.abort();
      let resolved = 0;
      const result = await prepareInlineImageResources(
        planned,
        async () => {
          resolved++;
          if (kind === 'resolver') throw Error('Resolver refused this asset.');
          return {
            bytes: kind === 'invalid-bytes' ? new Uint8Array([1, 2, 3]) : png,
            mime: kind === 'mime' ? 'image/jpeg' : 'image/png',
            widthPoints: kind === 'dimensions' ? -1 : 12,
            heightPoints: 8,
          };
        },
        kind === 'decode'
          ? {
              async decode() {
                return { pixelWidth: 2, pixelHeight: 1, dpiX: 96, dpiY: 96 };
              },
            }
          : decode,
        abort.signal
      );
      expect(result.success).toBe(false);
      expect(resolved).toBe(kind === 'abort' ? 0 : 1);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(media(editor.surface!.session.currentPackage())).toEqual([]);
    }
  } finally {
    editor.destroy();
  }
});

test('a source revision change during image resolution refuses stale execution without adding an undo entry', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    const planned = await plan(editor, 'Keep. A ![Cat](asset:cat) Z');
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const preparing = prepareInlineImageResources(
      planned,
      async () => {
        await gate;
        return { bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 };
      },
      decode
    );
    const paragraphId = editor.surface!.session.paragraphIds()[0]!;
    expect(
      editor.surface!.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId, offset: 0, text: '!' }],
        { kind: 'body' },
        [],
        false,
        true
      ).rejected
    ).toBe(false);
    const before = editor.surface!.session.currentPackage();
    release();
    const prepared = await preparing;
    if (!prepared.success) throw Error(prepared.error);
    expect(applyDirectPlan(editor, planned, { ...context, images: prepared.images })).toMatchObject(
      { success: false, code: 'stale-revision' }
    );
    expect(editor.surface!.session.currentPackage()).toBe(before);
    expect(media(before)).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

for (const kind of ['heading', 'table'] as const)
  test(`inline image insertion retains existing ${kind} paragraphs and their native structure`, async () => {
    const source =
      kind === 'heading'
        ? docxFixture(
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Keep.</w:t></w:r></w:p>',
            {
              styles:
                '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>',
            }
          )
        : docxFixture(
            '<w:p><w:r><w:t>Outside</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>H</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Keep.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
          );
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: source,
      imageDecodePort: decode,
    });
    try {
      const ids = editor.surface!.session.paragraphIds();
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const view = await readWorkingView(editor.surface!.session, revision);
      if (!view.success) throw Error(view.error);
      const target = view.view.projection.text.replace('Keep', 'Keep ![Cell](asset:cell "Photo")');
      const planned = await plan(editor, target);
      expect(planned.operations.some((edit) => edit.kind === 'inline-atoms')).toBe(true);
      expect(
        planned.operations.some((edit) => edit.kind === 'siblings' || edit.kind === 'table')
      ).toBe(false);
      const prepared = await prepareInlineImageResources(
        planned,
        async () => ({ bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 }),
        decode
      );
      if (!prepared.success) throw Error(prepared.error);
      const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
      if (!result.success) throw Error(result.error);
      expect(editor.surface!.session.paragraphIds()).toEqual(ids);
      expect((await plan(editor, target)).operations).toEqual([]);
      const pkg = editor.surface!.session.currentPackage();
      expect(buildBlockIndex(pkg).diagnostics).toEqual([]);
      expect(media(pkg)).toHaveLength(1);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(media(editor.surface!.session.currentPackage())).toEqual([]);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      expect(media(loaded.package)).toEqual(media(pkg));
    } finally {
      editor.destroy();
    }
  });

test('unsupported marked and linked image ownership refuses before resolving resources', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    const before = editor.surface!.session.currentPackage();
    for (const target of [
      'Keep. **![Cat](asset:cat)** Z',
      'Keep. [![Cat](asset:cat)](https://example.test) Z',
    ]) {
      const planned = await plan(editor, target);
      let resolved = false;
      const result = await prepareInlineImageResources(
        planned,
        async () => {
          resolved = true;
          return { bytes: png, mime: 'image/png' };
        },
        decode
      );
      expect(result).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(resolved).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
    }
  } finally {
    editor.destroy();
  }
});

test('images in nonstandard main parts resolve their actual native media through DOCX reopening', async () => {
  const entries = unzipSync(docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'));
  entries['custom/main.xml'] = entries['word/document.xml']!;
  entries['custom/_rels/main.xml.rels'] = entries['word/_rels/document.xml.rels']!;
  delete entries['word/document.xml'];
  delete entries['word/_rels/document.xml.rels'];
  entries['_rels/.rels'] = strToU8(
    strFromU8(entries['_rels/.rels']!).replace('word/document.xml', 'custom/main.xml')
  );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace('/word/document.xml', '/custom/main.xml')
  );
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: zipSync(entries),
    imageDecodePort: decode,
  });
  try {
    const target = 'Keep. A ![Cat](asset:cat) Z';
    const planned = await plan(editor, target);
    const prepared = await prepareInlineImageResources(
      planned,
      async () => ({ bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 }),
      decode
    );
    if (!prepared.success) throw Error(prepared.error);
    const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
    if (!result.success) throw Error(result.error);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw Error(saved.reason);
    for (const pkg of [editor.surface!.session.currentPackage(), saved.package]) {
      expect(pkg.mainDocumentPart).toBe('/custom/main.xml');
      const image = [...buildBlockIndex(pkg).byId.values()].find(
        (block) => block.payload.type === 'inline_image'
      )!;
      const record = (pkg.relationships.get(pkg.mainDocumentPart) ?? []).find((entry) =>
        entry.type.endsWith('/image')
      )!;
      expect(record).toBeDefined();
      const resource = resolveImageRelationship([record], pkg.mainDocumentPart, record.id);
      expect(resource.mode).toBe('internal');
      if (resource.mode !== 'internal') throw Error('The image has no native media relationship.');
      expect(resource.partName).toBe(image.payload.attributes.mediaPartName);
      expect(pkg.partBytes.get(resource.partName)).toEqual(png);
    }
    expect((await plan(editor, target)).operations).toEqual([]);
  } finally {
    editor.destroy();
  }
});
