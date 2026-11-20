import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { MAX_NUMBERING_DEFINITIONS } from '@docx-editor.dev/core/layout';
import {
  DEFAULT_OOXML_PACKAGE_LIMITS,
  WML_NAMESPACE_URI,
  authorableHyperlinkTarget,
  escapeXmlAttribute,
  isValidXmlText,
  readZip,
  sanitizeHref,
} from '@docx-editor.dev/core/store';

import { type DocumentNode } from './document-types';
import { decodeFragment, HeadingFragments } from './document-word-headings';

export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const MARKDOWN = 'urn:docx-editor:markdown:format:1';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const DOCUMENT = '/word/document.xml';
const RELS = '/word/_rels/document.xml.rels';
const STYLES = '/word/styles.xml';
const TYPES = '/[Content_Types].xml';

export const MAX_XML_LENGTH = 8 * 1024 * 1024;

export function attribute(value: unknown): string {
  if (typeof value !== 'string' || !isValidXmlText(value)) {
    throw new Error('Invalid Markdown XML attribute');
  }
  return escapeXmlAttribute(value);
}

export function formatAttribute(key: string, value: unknown): string {
  return value == null ? '' : ` md:${key}="${attribute(value)}"`;
}

export function paragraph(content: string, properties = '', attrs = ''): string {
  return `<w:p${attrs}>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${content}</w:p>`;
}

export function style(id: string): string {
  return `<w:pStyle w:val="${id}"/>`;
}

export function control(tag: string, content: string, attrs = '', properties = ''): string {
  return `<w:sdt${attrs}><w:sdtPr><w:tag w:val="markdown:${tag}"/>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

const MONOSPACE = '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr>';
const CODE_STYLE =
  '<w:style w:type="character" w:styleId="HTMLCode"><w:name w:val="HTML Code"/>' +
  MONOSPACE +
  '</w:style>' +
  '<w:style w:type="paragraph" w:styleId="SourceCode"><w:name w:val="Source Code"/>' +
  '<w:basedOn w:val="NoSpacing"/><w:pPr><w:shd w:val="clear" w:fill="F2F2F2"/></w:pPr>' +
  MONOSPACE +
  '</w:style>';
const LINK_STYLE =
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
  '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>';
const TABLE_STYLE =
  '<w:style w:type="table" w:styleId="MarkdownTable"><w:name w:val="Markdown Table"/>' +
  '<w:basedOn w:val="TableNormal"/><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr>' +
  '<w:tcPr><w:shd w:val="clear" w:fill="E8E8E8"/></w:tcPr></w:tblStylePr></w:style>';

export interface ListNumbering {
  id: number;
  level: number;
}

export class WordParts {
  private readonly links = new Map<string, string>();
  private readonly lists: string[] = [];

  private readonly bookmarks = new Map<string, number>();
  private readonly headingIds = new WeakMap<DocumentNode, number>();

  prepareBookmarks(headings: DocumentNode[], targets: Set<string>): void {
    if (!targets.size) return;
    const fragments = new HeadingFragments();
    for (const node of headings) {
      const fragment = fragments.next(node);
      if (!targets.has(fragment)) continue;
      const id = this.bookmarks.size + 1;
      this.bookmarks.set(fragment, id);
      this.headingIds.set(node, id);
    }
  }

  headingContent(node: DocumentNode, runs: string): string {
    const id = this.headingIds.get(node);
    if (!id) return runs;
    return `<w:bookmarkStart w:id="${id}" w:name="MarkdownHeading${id}" md:heading="1"/>${runs}<w:bookmarkEnd w:id="${id}"/>`;
  }

  hyperlink(raw: unknown): string {
    if (typeof raw !== 'string') throw new Error('Missing Markdown link destination');
    const target = sanitizeHref(raw);
    if (!target.ok) throw new Error('Unsupported Markdown link destination');
    if (target.href.startsWith('#')) {
      const fragment = target.href.slice(1);
      const id = this.bookmarks.get(decodeFragment(fragment));
      return id
        ? ` w:anchor="MarkdownHeading${id}"` + formatAttribute('fragment', fragment)
        : ` w:anchor="${attribute(fragment)}"`;
    }
    // A relative Markdown destination has no base URI in a fresh DOCX. Keep its exact
    // destination for export without resolving it against the host that opens the document.
    if (!/^[a-z][a-z0-9+.-]*:/i.test(target.href)) return formatAttribute('relative', target.href);
    if (!authorableHyperlinkTarget(target.href))
      throw new Error('Unsupported Markdown link destination');
    let id = this.links.get(target.href);
    if (!id) {
      if (this.links.size + 4 >= DEFAULT_OOXML_PACKAGE_LIMITS.maxRelationships) {
        throw new Error('Too many Markdown links');
      }
      id = `rIdMarkdownLink${this.links.size + 1}`;
      this.links.set(target.href, id);
    }
    return ` r:id="${id}"`;
  }

  numbering(ordered: boolean, start: unknown, level: number): ListNumbering {
    if (level > 8 || this.lists.length >= MAX_NUMBERING_DEFINITIONS) {
      throw new Error('Markdown list exceeds native numbering limits');
    }
    if (typeof start !== 'number' || !Number.isInteger(start) || start < 0 || start > 9999) {
      throw new Error('Unsupported Markdown list start');
    }
    const id = this.lists.length + 1;
    const indent = (level + 1) * 720;
    this.lists.push(
      `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="multilevel"/>` +
        `<w:lvl w:ilvl="${level}"><w:start w:val="${start}"/><w:numFmt w:val="${ordered ? 'decimal' : 'bullet'}"/>` +
        `<w:lvlText w:val="${ordered ? `%${level + 1}.` : '•'}"/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:tabs><w:tab w:val="num" w:pos="${indent}"/></w:tabs>` +
        `<w:ind w:left="${indent}" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`
    );
    return { id, level };
  }

  finish(body: string): Map<string, Uint8Array> {
    const blank = readZip(blankDocumentBytes());
    if (!blank.ok) throw new Error(blank.reason);
    const entries = new Map(blank.entries);
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    function insert(part: string, before: string, xml: string) {
      const original = decoder.decode(entries.get(part));
      if (!original.includes(before)) throw new Error(`Missing blank document element in ${part}`);
      entries.set(part, encoder.encode(original.replace(before, () => xml + before)));
    }
    const preserved =
      'heading fragment soft marks outerMarks relative src title sourceLabel identifier label referenceType lang meta url emptyRoot listMarker listSpread itemSpread numId'
        .split(' ')
        .map((key) => 'md:' + key)
        .join(' ');
    const ns = ` xmlns:r="${R}" xmlns:w14="${W14}" xmlns:md="${MARKDOWN}" xmlns:mc="${MC}" mc:Ignorable="md w14" mc:PreserveAttributes="${preserved}"`;
    const template = decoder.decode(entries.get(DOCUMENT));
    entries.set(
      DOCUMENT,
      encoder.encode(
        template.replace('<w:document ', `<w:document${ns} `).replace('<w:p/>', () => body)
      )
    );
    insert(STYLES, '</w:styles>', CODE_STYLE + LINK_STYLE + TABLE_STYLE);

    const relationships: string[] = [];
    if (this.lists.length) {
      const instances = this.lists.map(
        (_, i) => `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${i + 1}"/></w:num>`
      );
      entries.set(
        '/word/numbering.xml',
        encoder.encode(
          `<w:numbering xmlns:w="${WML_NAMESPACE_URI}">${this.lists.join('')}${instances.join('')}</w:numbering>`
        )
      );
      insert(
        TYPES,
        '</Types>',
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      );
      relationships.push(
        `<Relationship Id="rIdMarkdownNumbering" Type="${R}/numbering" Target="numbering.xml"/>`
      );
    }
    for (const [target, id] of this.links) {
      relationships.push(
        `<Relationship Id="${id}" Type="${R}/hyperlink" Target="${attribute(target)}" TargetMode="External"/>`
      );
    }
    insert(RELS, '</Relationships>', relationships.join(''));
    return entries;
  }
}
