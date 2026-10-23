import type { InlineImageRuntime } from '../native/inline-image-runtime';
import { escapeXml } from '@docx-editor.dev/core/store';
import { blockTag } from '../blocks/payload-opc-codec';
import type { InsertedInlineAtom } from './inserted-atoms';
import { SCIOBOT_CODE_STYLE_ID } from '../native/inline-code';

/** New control fallback only. Character styles are applied after native paragraph allocation. */
export function inlineAtomXml(
  op: InsertedInlineAtom,
  images?: InlineImageRuntime,
  codeStyle = false
): { readonly xml: string; readonly length: number } | null {
  if (op.kind === 'drawing') return images?.render(op) ?? null;
  if (op.kind !== 'inline-control') return null;
  const properties =
    (codeStyle && op.marks.includes('code') ? `<w:rStyle w:val="${SCIOBOT_CODE_STYLE_ID}"/>` : '') +
    op.marks
      .filter((mark) => mark !== 'code')
      .map((mark) => (mark === 'bold' ? '<w:b/>' : mark === 'italic' ? '<w:i/>' : '<w:strike/>'))
      .join('');
  const run = (text: string) =>
    `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}${text}</w:r>`;
  const content = op.fallback
    .split(/(\n|\t|\u001e|\u001f)/)
    .filter(Boolean)
    .map((value) =>
      run(
        value === '\n'
          ? '<w:br/>'
          : value === '\t'
            ? '<w:tab/>'
            : value === '\u001e'
              ? '<w:noBreakHyphen/>'
              : value === '\u001f'
                ? '<w:softHyphen/>'
                : `<w:t xml:space="preserve">${escapeXml(value)}</w:t>`
      )
    )
    .join('');
  return {
    length: op.fallback.length,
    xml: `<w:sdt><w:sdtPr><w:alias w:val="${escapeXml(op.label)}"/><w:tag w:val="${escapeXml(blockTag(op.payload.id))}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`,
  };
}
export const inlineAtomKey = (location: readonly number[], path: readonly number[]) =>
  location.join('/') + ':' + path.join('/');
export function inlineAtomsByPath(atoms: readonly InsertedInlineAtom[]) {
  return new Map(
    atoms
      .filter((op) => op.location.kind === 'new-paragraph')
      .map((op) => [
        inlineAtomKey(
          (op.location as Extract<typeof op.location, { kind: 'new-paragraph' }>).path,
          op.path
        ),
        op,
      ])
  );
}
