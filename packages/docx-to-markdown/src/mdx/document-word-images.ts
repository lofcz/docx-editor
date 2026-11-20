import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { geometryOfSection, readSectionProperties } from '@docx-editor.dev/core/layout';
import {
  escapeXmlAttribute,
  isValidXmlText,
  readOoxmlPackage,
  readZip,
  relsPartNameFor,
  resolveContentTypeOf,
  resolveImageRelationship,
  resolveImageResourceLimits,
  serializeOoxmlPart,
  validateRasterHeader,
  withEmbeddedImage,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { DocumentNode } from './document-types';

const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_HEADER = /^data:(image\/(?:png|jpeg));base64,/i;
const MAX_IMAGE_BYTES = 1024 * 1024;
// Each distinct image allocates a part and an owner relationship in the blank package.
const MAX_MEDIA_PARTS = 128;
type Mime = 'image/png' | 'image/jpeg';

// Find the orientation SHORT in the primary TIFF directory of an EXIF APP1 segment.
function tiffOrientationByte(view: DataView, tiff: number, end: number): number | undefined {
  if (tiff + 8 > end) return undefined;
  const order = view.getUint16(tiff);
  const le = order === 0x4949;
  if ((!le && order !== 0x4d4d) || view.getUint16(tiff + 2, le) !== 42) return undefined;
  const ifd = tiff + view.getUint32(tiff + 4, le);
  if (ifd < tiff + 8 || ifd + 2 > end) return undefined;
  const count = view.getUint16(ifd, le);
  const last = ifd + 2 + 12 * count;
  if (count > 4096 || last > end) return undefined;
  for (let entry = ifd + 2; entry < last; entry += 12) {
    if (view.getUint16(entry, le) !== 274) continue;
    if (view.getUint16(entry + 2, le) !== 3 || view.getUint32(entry + 4, le) !== 1)
      return undefined;
    const value = view.getUint16(entry + 8, le);
    return value >= 1 && value <= 8 ? entry + (le ? 8 : 9) : undefined;
  }
  return undefined;
}

// Word applies DrawingML transforms. Browsers also apply EXIF. Set orientation to 1
// in the media, and restore the original byte on MDX export.
function jpegOrientationByte(bytes: Uint8Array): number | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let field: number | undefined;
  for (let i = 2; i + 4 <= bytes.length && bytes[i] === 0xff;) {
    const marker = bytes[++i]!;
    if (marker === 0xff) continue;
    i++;
    if (marker === 0xd9 || marker === 0x00) break;
    if (marker === 1 || (marker >= 0xd0 && marker <= 0xd8)) continue;
    const end = i + view.getUint16(i);
    if (end < i + 2 || end > bytes.length) break;
    if (marker === 0xda) {
      const components = bytes[i + 2];
      if (
        components &&
        end === i + 6 + 2 * components &&
        end < bytes.length - 2 &&
        bytes.at(-2) === 0xff &&
        bytes.at(-1) === 0xd9
      )
        return field;
      break;
    }
    const tiff = i + 8;
    if (
      marker === 0xe1 &&
      tiff <= end &&
      view.getUint32(i + 2) === 0x45786966 &&
      view.getUint16(i + 6) === 0
    ) {
      field ??= tiffOrientationByte(view, tiff, end);
    }
    i = end;
  }
  throw new Error('Invalid Markdown image bytes for image/jpeg');
}

const EXIF_TRANSFORMS = [
  '',
  '',
  ' flipH="1"',
  ' rot="10800000"',
  ' flipV="1"',
  ' rot="5400000" flipV="1"',
  ' rot="5400000"',
  ' rot="5400000" flipH="1"',
  ' rot="16200000"',
];

function decode(source: string, prefix: string): Uint8Array {
  const length = source.length - prefix.length;
  const padding = source.endsWith('==') ? 2 : source.endsWith('=') ? 1 : 0;
  if (length * 0.75 - padding > MAX_IMAGE_BYTES)
    throw new RangeError('Markdown image exceeds the 1 MiB limit');
  if (!length || length % 4) throw new Error('Invalid Markdown image base64');
  const payload = source.slice(prefix.length);
  let binary: string;
  try {
    binary = atob(payload);
  } catch {
    throw new Error('Invalid Markdown image base64');
  }
  // atob accepts whitespace and unused nonzero padding bits. Require a reversible spelling.
  if (btoa(binary) !== payload) throw new Error('Invalid Markdown image base64');
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < bytes.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function hasPngImageData(bytes: Uint8Array): boolean {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let hasData = false;
  for (let offset = 8; offset + 12 <= bytes.length;) {
    const size = view.getUint32(offset);
    const type = view.getUint32(offset + 4);
    if (size > bytes.length - offset - 12) return false;
    offset += 12 + size;
    if (type === 0x49444154 && size > 0) hasData = true; // IDAT
    if (type === 0x49454e44) return hasData && size === 0 && offset === bytes.length; // IEND
  }
  return false;
}

function raster(bytes: Uint8Array, mime: Mime) {
  const header = validateRasterHeader(bytes, mime);
  if (!header || (mime === 'image/png' && !hasPngImageData(bytes)))
    throw new Error('Invalid Markdown image bytes for ' + mime);
  const limits = resolveImageResourceLimits();
  if (
    header.pixelWidth > limits.maxDimension ||
    header.pixelHeight > limits.maxDimension ||
    header.pixelWidth * header.pixelHeight > limits.maxPixels
  )
    throw new RangeError('Markdown image dimensions exceed the raster limit');
  return header;
}

interface Media {
  prefix: string;
  id: string;
  width: number;
  height: number;
  orientation: number;
}

function sourceAttributes(media: Media): string {
  return (
    ` md:src="${media.prefix}"` +
    (media.orientation > 1 ? ` md:orientation="${media.orientation}"` : '')
  );
}

export class WordImages {
  readonly referenceIds = new Set<string>();
  private readonly media = new Map<string, Media>();
  private pkg?: OoxmlPackage;
  private drawingId = 0;

  private embed(source: string): Media | undefined {
    if (!/^data:/i.test(source)) return undefined;
    const saved = this.media.get(source);
    if (saved) return saved;
    const match = DATA_HEADER.exec(source);
    if (!match) throw new Error('Markdown image requires a base64 PNG or JPEG data URI');
    if (this.media.size >= MAX_MEDIA_PARTS)
      throw new RangeError('Too many embedded Markdown images');
    const mime = match[1]!.toLowerCase() as Mime;
    const bytes = decode(source, match[0]);
    const field = mime === 'image/jpeg' ? jpegOrientationByte(bytes) : undefined;
    const orientation = field === undefined ? 1 : bytes[field]!;
    if (field !== undefined) bytes[field] = 1;
    const header = raster(bytes, mime);
    if (!this.pkg) {
      const blank = readOoxmlPackage(blankDocumentBytes());
      if (!blank.ok) throw new Error(blank.reason);
      this.pkg = blank.package;
    }
    const part = withEmbeddedImage(this.pkg, this.pkg.mainDocumentPart, { bytes, mime });
    if (!part.ok) throw new Error('Cannot embed Markdown image: ' + part.reason);
    this.pkg = part.pkg;
    const page = geometryOfSection(
      readSectionProperties(this.pkg.parts.get(this.pkg.mainDocumentPart)!)
    );
    const width = (header.pixelWidth * 72) / (header.dpiX ?? 96);
    const height = (header.pixelHeight * 72) / (header.dpiY ?? 96);
    const rotated = orientation >= 5;
    const scale = Math.min(
      1,
      (page.width - page.margin.left - page.margin.right) / (rotated ? height : width),
      (page.height - page.margin.top - page.margin.bottom) / (rotated ? width : height)
    );
    const media = {
      prefix: match[0],
      id: part.relationshipId,
      width: width * scale,
      height: height * scale,
      orientation,
    };
    this.media.set(source, media);
    return media;
  }

  drawing(
    source: string,
    alt: unknown,
    title: unknown,
    reference: string,
    context: { imageLinkId?: string; imageWidthPoints?: number }
  ): string | undefined {
    const media = this.embed(source);
    if (!media) return undefined;
    if (
      typeof alt !== 'string' ||
      !isValidXmlText(alt) ||
      (title != null && (typeof title !== 'string' || !isValidXmlText(title)))
    ) {
      throw new Error('Invalid Markdown image label');
    }
    const titleAttr = typeof title === 'string' ? ` title="${escapeXmlAttribute(title)}"` : '';
    const width = media.orientation >= 5 ? media.height : media.width;
    const scale = Math.min(1, (context.imageWidthPoints ?? width) / width);
    const cx = Math.max(1, Math.round(media.width * scale * 12700));
    const cy = Math.max(1, Math.round(media.height * scale * 12700));
    const extent = `cx="${cx}" cy="${cy}"`;
    const id = ++this.drawingId;
    const click = context.imageLinkId
      ? `<a:hlinkClick r:id="${escapeXmlAttribute(context.imageLinkId)}"/>`
      : '';
    return (
      `<w:drawing${sourceAttributes(media)}${reference} xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}">` +
      `<wp:inline><wp:extent ${extent}/><wp:docPr id="${id}" name="Picture ${id}" descr="${escapeXmlAttribute(alt)}"${titleAttr}>${click}</wp:docPr>` +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      `<a:graphic><a:graphicData uri="${PIC}"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="0" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${media.id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm${EXIF_TRANSFORMS[media.orientation]}><a:off x="0" y="0"/><a:ext ${extent}/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    );
  }

  definition(node: DocumentNode): string | undefined {
    if (
      typeof node.identifier !== 'string' ||
      !this.referenceIds.has(node.identifier) ||
      typeof node.url !== 'string'
    )
      return undefined;
    const media = this.embed(node.url);
    return media && sourceAttributes(media) + ` md:imageId="${media.id}"`;
  }

  templateEntries(): Map<string, Uint8Array> {
    if (!this.pkg) {
      const blank = readZip(blankDocumentBytes());
      if (!blank.ok) throw new Error(blank.reason);
      return new Map(blank.entries);
    }
    // withEmbeddedImage writes media and content types as bytes, and updates the owner's rels tree.
    const entries = new Map(this.pkg.partBytes);
    const rels = relsPartNameFor(this.pkg.mainDocumentPart);
    entries.set(rels, new TextEncoder().encode(serializeOoxmlPart(this.pkg.parts.get(rels)!)));
    return entries;
  }
}

export class WordImageReader {
  private readonly sources = new Map<string, string>();
  constructor(
    private readonly pkg: OoxmlPackage,
    private remaining: number
  ) {}

  source(id: string, prefix: string | undefined, originalOrientation: string | undefined): string {
    if (!prefix || prefix.length > 23 || DATA_HEADER.exec(prefix)?.[0] !== prefix)
      throw new Error('Invalid stored Markdown image data URI');
    if (originalOrientation !== undefined && !/^[2-8]$/.test(originalOrientation))
      throw new Error('Invalid stored Markdown image orientation');
    const key = [id, prefix, originalOrientation].join('\0');
    const saved = this.sources.get(key);
    if (saved) return saved;
    const owner = this.pkg.mainDocumentPart;
    const rel = resolveImageRelationship(this.pkg.relationships.get(owner), owner, id);
    if (rel.mode !== 'internal')
      throw new Error('Expected an internal Markdown image relationship');
    const bytes = this.pkg.partBytes.get(rel.partName);
    const mime = resolveContentTypeOf(this.pkg, rel.partName);
    if (!bytes || (mime !== 'image/png' && mime !== 'image/jpeg'))
      throw new Error('Missing or unsupported Markdown image media');
    if (bytes.length > MAX_IMAGE_BYTES)
      throw new RangeError('Markdown image exceeds the 1 MiB limit');
    this.remaining -= 4 * Math.ceil(bytes.length / 3) + prefix.length;
    if (this.remaining < 0)
      throw new RangeError('Embedded Markdown images exceed the document size limit');
    raster(bytes, mime);
    const field = mime === 'image/jpeg' ? jpegOrientationByte(bytes) : undefined;
    let original = bytes;
    if (originalOrientation && field !== undefined && bytes[field] === 1) {
      original = new Uint8Array(bytes);
      original[field] = Number(originalOrientation);
    }
    const chunks: string[] = [];
    for (let i = 0; i < original.length; i += 8192) {
      chunks.push(String.fromCharCode(...original.subarray(i, i + 8192)));
    }
    const canonical = `data:${mime};base64,`;
    const url = (prefix.toLowerCase() === canonical ? prefix : canonical) + btoa(chunks.join(''));
    this.sources.set(key, url);
    return url;
  }

  drawing(
    node: OoxmlElement,
    prefix: string | undefined,
    orientation: string | undefined
  ): DocumentNode & { url: string } {
    if (node.kind !== 'drawing' || !prefix)
      throw new Error('Unsupported Word run content: drawing');
    const inline = node.children.find((n) => n.kind === 'inlineDrawing');
    const props = inline?.children.find((n) => n.kind === 'drawingDocPr');
    const graphic = inline?.children.find((n) => n.kind === 'drawingGraphic');
    const data = graphic?.children.find((n) => n.kind === 'drawingGraphicData');
    const pic = data?.children.find((n) => n.kind === 'picture');
    const fill = pic?.children.find((n) => n.kind === 'pictureBlipFill');
    const blip = fill?.children.find((n) => n.kind === 'pictureBlip');
    const id = blip?.attributes.find((a) => a.namespaceUri === R && a.localName === 'embed')?.value;
    if (!id || !props) throw new Error('Invalid embedded Markdown picture');
    const prop = (name: string) =>
      props.attributes.find((a) => !a.namespaceUri && a.localName === name)?.value;
    return {
      type: 'image',
      url: this.source(id, prefix, orientation),
      alt: prop('descr') ?? '',
      title: prop('title') ?? null,
    };
  }
}
