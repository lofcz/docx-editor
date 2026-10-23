import {
  escapeXml,
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import {
  compileInlineTarget,
  type InlineTargetHyperlink,
  type InlineTargetToken,
} from '../diff/inline-target';
import type { TargetAstNode } from '../projection/parse-target';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const allowed = (node: TargetAstNode) =>
  Object.keys(node).every(
    (key) =>
      ['type', 'children', 'position', ...(node.type === 'heading' ? ['depth'] : [])].includes(
        key
      ) ||
      (key === 'data' &&
        node.data !== null &&
        typeof node.data === 'object' &&
        !Array.isArray(node.data) &&
        Object.keys(node.data).length === 0)
  );

/** An inline code range in a new paragraph, in model offsets of its own text. */
export interface NewParagraphCode {
  readonly paragraphIndex: number;
  readonly start: number;
  readonly end: number;
}

export interface NewParagraphLink {
  readonly paragraphIndex: number;
  readonly start: number;
  readonly end: number;
  readonly target: InlineTargetHyperlink['target'];
  readonly title: string | null;
}

/** Materialize NEW AST paragraphs only. Imported canonical paragraphs never enter this writer. */
export function newParagraphSeeds(
  nodes: readonly TargetAstNode[],
  options: {
    /**
     * Report inline code ranges instead of refusing them. Seeds cannot carry `w:rStyle`, so the
     * caller applies the application character style (N18) after allocation.
     */
    readonly code?: boolean;
    /** Validated new controls only; no resource resolution occurs in this seed builder. */
    readonly atom?: (
      token: Extract<InlineTargetToken, { kind: 'atom' }>,
      paragraphIndex: number
    ) => { readonly xml: string; readonly length: number } | null;
  } = {}
):
  | {
      success: true;
      paragraphs: readonly OoxmlParagraphNode[];
      links: readonly NewParagraphLink[];
      codes: readonly NewParagraphCode[];
    }
  | NativeEditFailure {
  if (!nodes.length) return { success: true, paragraphs: [], links: [], codes: [] };
  const paragraphs: string[] = [];
  const links: NewParagraphLink[] = [];
  const codes: NewParagraphCode[] = [];
  let length = 0;
  for (const node of nodes) {
    if (!['paragraph', 'heading'].includes(node.type) || !allowed(node))
      return fail('Inserted block types and metadata need their own native fragment executor.');
    if (
      node.type === 'heading' &&
      (!Number.isInteger(node.depth) || Number(node.depth) < 1 || Number(node.depth) > 6)
    )
      return fail('A heading depth must be between one and six.');
    const inline = compileInlineTarget(node);
    if (!inline.success) return inline;
    if (
      (inline.target.plainText === null && !options.atom) ||
      (!options.code && inline.target.tokens.some((token) => token.marks.includes('code')))
    )
      return fail('Inline atoms and code need their own native fragment representation.');
    const rendered = new Map<
      InlineTargetToken,
      { readonly xml: string; readonly length: number }
    >();
    for (const token of inline.target.tokens)
      if (token.kind === 'atom') {
        const atom = options.atom?.(token, paragraphs.length);
        if (!atom || !Number.isSafeInteger(atom.length) || atom.length < 1)
          return fail(
            'Inserted atoms need their supported native payload or resource representation.'
          );
        rendered.set(token, atom);
      }
    let offset = 0;
    const spans = new Map<string, { start: number; end: number }>();
    for (const token of inline.target.tokens) {
      const width = token.kind === 'text' ? token.value.length : rendered.get(token)!.length;
      if (token.hyperlink !== null) {
        const key = JSON.stringify(token.hyperlink),
          previous = spans.get(key);
        spans.set(key, { start: previous?.start ?? offset, end: offset + width });
      }
      if (token.marks.includes('code') && width) {
        const previous = codes.at(-1);
        if (previous?.paragraphIndex === paragraphs.length && previous.end === offset)
          codes[codes.length - 1] = { ...previous, end: offset + width };
        else
          codes.push({
            paragraphIndex: paragraphs.length,
            start: offset,
            end: offset + width,
          });
      }
      offset += width;
    }
    for (const link of inline.target.hyperlinks) {
      const span = spans.get(JSON.stringify(link.path));
      if (!span || span.start === span.end)
        return fail('A native hyperlink needs nonempty owned text.');
      links.push({
        paragraphIndex: paragraphs.length,
        ...span,
        target: link.target,
        title: link.title,
      });
    }
    const runs: string[] = [];
    for (const token of inline.target.tokens) {
      if (token.kind === 'atom') {
        runs.push(rendered.get(token)!.xml);
        continue;
      }
      const properties = token.marks
        .filter((mark) => mark !== 'code')
        .map((mark) => (mark === 'bold' ? '<w:b/>' : mark === 'italic' ? '<w:i/>' : '<w:strike/>'))
        .join('');
      for (const value of token.value.split(/(\n|\t|\u001e|\u001f)/)) {
        if (!value) continue;
        const content =
          value === '\n'
            ? '<w:br/>'
            : value === '\t'
              ? '<w:tab/>'
              : value === '\u001e'
                ? '<w:noBreakHyphen/>'
                : value === '\u001f'
                  ? '<w:softHyphen/>'
                  : `<w:t xml:space="preserve">${escapeXml(value)}</w:t>`;
        runs.push(`<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}${content}</w:r>`);
      }
    }
    const xml = `<w:p>${node.type === 'heading' ? `<w:pPr><w:outlineLvl w:val="${Number(node.depth) - 1}"/></w:pPr>` : ''}${runs.join('')}</w:p>`;
    length += xml.length;
    if (length > 8 * 1024 * 1024)
      return fail('New paragraph XML exceeds its bounded fragment budget.');
    paragraphs.push(xml);
  }
  const xml = `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${paragraphs.join('')}</w:body></w:document>`;
  if (new TextEncoder().encode(xml).byteLength > 8 * 1024 * 1024)
    return fail('New paragraph XML exceeds its byte budget.');
  const parsed = readOoxmlPart(xml, {
    name: '/sciobot-new-paragraphs.xml',
    contentType: 'application/xml',
  });
  if (!parsed.ok) return fail(`New paragraph XML refused: ${parsed.reason}.`);
  const body = parsed.part.root.children[0];
  if (!body || body.kind === 'textValue' || body.children.some((node) => node.kind !== 'paragraph'))
    return fail('New paragraphs did not produce complete native blocks.');
  return {
    success: true,
    paragraphs: body.children as readonly OoxmlParagraphNode[],
    links,
    codes,
  };
}
