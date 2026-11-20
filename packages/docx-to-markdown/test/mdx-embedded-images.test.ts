import { expect, test } from 'bun:test';
import {
  IMAGE_RELATIONSHIP_TYPE,
  TreeDocumentStore,
  inspectDrawingSupport,
  readOoxmlPackage,
  readZip,
  resolveImageRelationship,
  validatePackageInvariants,
  withPart,
  writeOoxmlPackage,
  writeZip,
} from '@docx-editor.dev/core/store';
import { importMdxDocument, exportMdxDocument } from '../src/mdx.ts';
import { parseDocumentSyntax } from '../src/mdx/document-syntax.ts';

// Two colored panels, 96 × 48 pixels. The JPEG stores 48 × 24 pixels with EXIF orientation 6.
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAGAAAAAwCAIAAABhdOiYAAAAhklEQVRo3u3QMRGAMADAQMAPOyuGUFQ5CGNmytgO/wpy2d/z2lby3GN2ws8xO2B1BgWDgkHBoGBQMCgYFAwKBgWDgkHBoGBQMCgYFAwKBgWDgkHBoGBQMCgYFAwKBgWDgkHBoGBQMCgYFAwKBgWDgkHBoGBQMCgYFAwKBgWDgkHBoGBQMCh8logC1E1Tm1MAAAAASUVORK5CYII=';
const JPEG =
  '/9j/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAYAAAAAAAD/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAAYADADASIAAhEBAxEB/8QAFgABAQEAAAAAAAAAAAAAAAAAAAUG/8QAFRABAQAAAAAAAAAAAAAAAAAAAAH/xAAXAQEBAQEAAAAAAAAAAAAAAAAHBgUI/8QAGREBAAMBAQAAAAAAAAAAAAAAAAZDgsHC/9oADAMBAAIRAxEAPwDHAMgrEUYnRRiuituej6dUa8gCvH6cAInQJFGArorbno+nVGvIArx+/9k=';
const png = 'data:image/png;base64,' + PNG;
const jpeg = 'data:image/jpeg;base64,' + JPEG;
const doc = '/word/document.xml';

function open(bytes: Uint8Array) {
  const result = readOoxmlPackage(bytes);
  if (!result.ok) throw new Error(result.reason);
  return result.package;
}

function editXml(bytes: Uint8Array, part: string, edit: (xml: string) => string) {
  const zip = readZip(bytes);
  if (!zip.ok) throw new Error(zip.reason);
  const entries = new Map(zip.entries);
  const xml = new TextDecoder().decode(entries.get(part));
  const changed = edit(xml);
  expect(changed).not.toBe(xml);
  entries.set(part, new TextEncoder().encode(changed));
  return writeZip(entries);
}

test('offline images use native inline pictures, links, dimensions, and original media after save', () => {
  const source = `A [**![Red & blue](${png} "Diagram")**](https://example.org/guide) and ![](${jpeg}).\n`;
  const saved = writeOoxmlPackage(open(importMdxDocument(source)));
  const pkg = open(saved);
  expect(validatePackageInvariants(pkg).ok).toBe(true);
  expect(
    inspectDrawingSupport(pkg).entries.map(({ state, graphicKind }) => [state, graphicKind])
  ).toEqual([
    ['projected', 'picture'],
    ['projected', 'picture'],
  ]);
  const rels = pkg.relationships.get(doc)!;
  const images = rels.filter((r) => r.type === IMAGE_RELATIONSHIP_TYPE);
  expect(images).toHaveLength(2);
  for (const [i, rel] of images.entries()) {
    const target = resolveImageRelationship(rels, doc, rel.id);
    if (target.mode !== 'internal') throw new Error('Expected an internal image');
    const bytes = pkg.partBytes.get(target.partName)!;
    const expected = Buffer.from(i === 0 ? PNG : JPEG, 'base64');
    // Only the TIFF orientation byte changes. The drawing applies the rotation once.
    if (i === 1) expected[31] = 1;
    expect(Buffer.from(bytes)).toEqual(expected);
  }
  expect(pkg.externalTargets.map((r) => r.rawTarget)).toEqual(['https://example.org/guide']);
  const xml = new TextDecoder().decode(pkg.partBytes.get(doc));
  expect(xml).toContain('descr="Red &amp; blue"');
  expect(xml).toContain('cx="914400" cy="457200"');
  expect(xml).toContain('cx="457200" cy="228600"');
  expect(xml).toContain('<a:xfrm rot="5400000">');
  expect(xml).toContain('<a:hlinkClick r:id="rIdMarkdownLink1"');
  expect(xml).not.toContain(PNG);
  expect(xml).not.toContain(JPEG);
  expect(parseDocumentSyntax(exportMdxDocument(saved))).toEqual(parseDocumentSyntax(source));
});

test('repeated image references share media and keep distinct native labels and drawing IDs', () => {
  const source = `[photo]: ${png.replace('image/png', 'IMAGE/PNG')} "Color"\n\n![First][photo] and *![Second][photo]*\n`;
  const pkg = open(importMdxDocument(source));
  const entries = inspectDrawingSupport(pkg).entries;
  expect(entries).toHaveLength(2);
  expect(
    pkg.relationships.get(doc)!.filter((r) => r.type === IMAGE_RELATIONSHIP_TYPE)
  ).toHaveLength(1);
  const xml = new TextDecoder().decode(pkg.partBytes.get(doc));
  expect(new Set([...xml.matchAll(/<wp:docPr [^>]*?\bid="(\d+)"/g)].map((m) => m[1])).size).toBe(2);
  expect(xml).not.toContain(PNG);
  expect(parseDocumentSyntax(exportMdxDocument(writeOoxmlPackage(pkg)))).toEqual(
    parseDocumentSyntax(source)
  );

  const store = new TreeDocumentStore(pkg.parts.get(doc)!);
  const changed = store.transact((ctx) =>
    ctx.apply({
      op: 'setDrawingMetadata',
      drawingNodeId: entries[0]!.drawingNodeId!,
      title: 'Color',
      description: 'New label',
    })
  );
  expect(changed.ok).toBe(true);
  const updated = writeOoxmlPackage(withPart(pkg, store.part));
  expect(parseDocumentSyntax(exportMdxDocument(updated))).toEqual(
    parseDocumentSyntax(source.replace('![First]', '![New label]'))
  );
});

test('physical PNG dimensions fit the printable page and the text area of a table cell', () => {
  const bytes = Buffer.from(PNG, 'base64');
  // A valid pHYs chunk at 400 pixels per meter makes this 96-pixel image 240 mm wide.
  const density = Buffer.from('000000097048597300000190000001900121f7eb7e', 'hex');
  const image = Buffer.concat([bytes.subarray(0, 33), density, bytes.subarray(33)]).toString(
    'base64'
  );
  const ref = '![Panels][print]';
  const source = `${ref}\n\n| View | Notes |\n| --- | --- |\n| ${ref} | Detail |\n\n[print]: data:image/png;base64,${image}\n`;
  const pkg = open(importMdxDocument(source));
  const xml = new TextDecoder().decode(pkg.partBytes.get(doc));
  expect(xml).toContain('<wp:extent cx="5943600" cy="2971800"/>'); // 468 × 234 pt
  expect(xml).toContain('<wp:extent cx="2834640" cy="1417320"/>'); // 223.2 × 111.6 pt
  expect(parseDocumentSyntax(exportMdxDocument(pkg))).toEqual(parseDocumentSyntax(source));
});

test('remote and relative images retain editable fallback text and inert destinations', () => {
  const source =
    '![Remote](https://example.org/photo.png "View") and ![](../relative/photo.jpg).\n';
  const pkg = open(importMdxDocument(source));
  expect(pkg.externalTargets).toHaveLength(0);
  expect(inspectDrawingSupport(pkg).entries).toHaveLength(0);
  expect(new TextDecoder().decode(pkg.partBytes.get(doc))).toContain('>Remote</w:t>');
  expect(parseDocumentSyntax(exportMdxDocument(writeOoxmlPackage(pkg)))).toEqual(
    parseDocumentSyntax(source)
  );
  // Previously saved documents kept the full data URI on their text content control.
  const legacy = editXml(importMdxDocument('![Old](photo.png)'), doc, (xml) =>
    xml.replace('md:src="photo.png"', `md:src="${png}"`)
  );
  expect(exportMdxDocument(legacy)).toBe(`![Old](${png})\n`);
});

test.each([
  ['data:image/png;base64,', /Invalid Markdown image base64/],
  ['data:image/png;base64,not*base64==', /Invalid Markdown image base64/],
  ['data:image/png;base64,AB==', /Invalid Markdown image base64/],
  ['data:image/svg+xml;base64,PHN2Zy8+', /base64 PNG or JPEG data URI/],
  ['data:image/png,abc', /base64 PNG or JPEG data URI/],
  ['data:image/jpeg;base64,' + PNG, /Invalid Markdown image bytes/],
  ['data:image/png;base64,' + PNG.slice(0, -12), /Invalid Markdown image bytes/],
  ['data:image/jpeg;base64,' + JPEG.slice(0, -4), /Invalid Markdown image bytes/],
] as const)('refuses malformed embedded image %#', (url, error) => {
  expect(() => importMdxDocument(`![image](${url})`)).toThrow(error);
});

test('refuses oversized decoded bytes and excessive raster dimensions', () => {
  const payload = 'A'.repeat(4 * Math.ceil((1024 * 1024 + 3) / 3));
  expect(() => importMdxDocument(`![image](data:image/png;base64,${payload})`)).toThrow(
    'Markdown image exceeds the 1 MiB limit'
  );
  const huge = Buffer.from(PNG, 'base64');
  huge.writeUInt32BE(0x7fffffff, 16);
  expect(() =>
    importMdxDocument(`![image](data:image/png;base64,${huge.toString('base64')})`)
  ).toThrow('dimensions exceed the raster limit');
});

test('export refuses an image relationship changed to an external target', () => {
  const altered = editXml(
    importMdxDocument(`![image](${png})`),
    '/word/_rels/document.xml.rels',
    (xml) =>
      xml.replace(
        /Target="media\/[^"]+"/,
        'TargetMode="External" Target="https://example.org/asset.png"'
      )
  );
  expect(() => exportMdxDocument(altered)).toThrow(
    'Expected an internal Markdown image relationship'
  );
});
