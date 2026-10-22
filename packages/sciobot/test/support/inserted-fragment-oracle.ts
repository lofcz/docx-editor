import {
  applyTreeOp,
  bodyStoryRoot,
  ensureHyperlinkRelationship,
  escapeXml,
  findNode,
  readOoxmlPart,
  removeNode,
  paragraphTextOf,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { blockTag, withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import type { InsertedCellPayloadBinding } from '../../src/diff/inserted-cell-fragments';
import type { TargetAstNode } from '../../src/projection/parse-target';
import { nodeChildren } from './native-table-oracle';
import { W } from './document';

/** Test-only codec for new fragments. This is not a general Markdown writer or live executor. */
export function materializeNewCellFragments(
  input: OoxmlPackage,
  partName: string,
  cells: readonly { readonly cellId?: string; readonly blocks: readonly TargetAstNode[] }[],
  payloads: readonly InsertedCellPayloadBinding[]
): OoxmlPackage {
  let pkg = input,
    part = input.parts.get(partName)!;
  for (const binding of payloads) {
    const stored = withPayloadRecord(pkg, binding.payload, binding.label);
    if (!stored.ok) throw new Error(stored.reason);
    pkg = stored.pkg;
  }
  const bindings = new Map(payloads.map((binding) => [binding.path.join('/'), binding]));
  const childAst = (node: TargetAstNode) => (node.children ?? []) as TargetAstNode[];
  const renderChildren = (node: TargetAstNode, path: number[], marks = ''): string =>
    childAst(node)
      .map((entry, index) => render(entry, [...path, index], marks))
      .join('');
  const run = (xml: string, marks: string) =>
    `<w:r>${marks ? `<w:rPr>${marks}</w:rPr>` : ''}${xml}</w:r>`;
  const render = (node: TargetAstNode, path: number[], marks = ''): string => {
    switch (node.type) {
      case 'text':
        return run(`<w:t xml:space="preserve">${escapeXml(String(node.value))}</w:t>`, marks);
      case 'strong':
        return renderChildren(node, path, `${marks}<w:b/>`);
      case 'emphasis':
        return renderChildren(node, path, `${marks}<w:i/>`);
      case 'delete':
        return renderChildren(node, path, `${marks}<w:strike/>`);
      case 'inlineCode':
        return run(
          `<w:t>${escapeXml(String(node.value))}</w:t>`,
          `${marks}<w:rStyle w:val="Code"/>`
        );
      case 'break':
        return run('<w:br/>', marks);
      case 'link': {
        const link = ensureHyperlinkRelationship(pkg, String(node.url), partName);
        if (!link) throw new Error('Unsafe fixture link');
        pkg = link.pkg;
        return `<w:hyperlink r:id="${escapeXml(link.relationshipId)}">${renderChildren(node, path, marks)}</w:hyperlink>`;
      }
      case 'paragraph':
        return `<w:p>${renderChildren(node, path)}</w:p>`;
      case 'table':
        return `<w:tbl><w:tblPr/><w:tblGrid>${childAst(childAst(node)[0]!)
          .map(() => '<w:gridCol w:w="1200"/>')
          .join('')}</w:tblGrid>${renderChildren(node, path)}</w:tbl>`;
      case 'tableRow':
        return `<w:tr><w:trPr/>${renderChildren(node, path)}</w:tr>`;
      case 'tableCell':
        return `<w:tc><w:tcPr/><w:p>${renderChildren(node, path)}</w:p></w:tc>`;
      case 'mdxJsxFlowElement': {
        const binding = bindings.get(path.join('/'));
        if (!binding) throw new Error('Missing fixture payload');
        const fallback = binding.payload.type.startsWith('shared_mdx_')
          ? `<w:p><w:r><w:t>${escapeXml(binding.label)}</w:t></w:r></w:p>`
          : renderChildren(node, path);
        return `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(binding.payload.id)}"/><w:alias w:val="${escapeXml(binding.label)}"/></w:sdtPr><w:sdtContent>${fallback}</w:sdtContent></w:sdt>`;
      }
      default:
        throw new Error(`Unsupported oracle fixture node: ${node.type}`);
    }
  };
  for (let index = 0; index < cells.length; index++) {
    const ast = cells[index]!.blocks;
    if (!ast.length) continue;
    const xml = ast.map((node, block) => render(node, [index, block])).join('');
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${xml}</w:body></w:document>`,
      { name: part.name, contentType: part.contentType }
    );
    if (!parsed.ok) throw new Error(parsed.reason);
    const seed = nodeChildren(findNode(part, cells[index]!.cellId!)!).find(
      (node) => node.kind === 'paragraph'
    )!;
    const pasted = applyTreeOp(part, {
      op: 'insertFragment',
      paragraphId: seed.id,
      offset: 0,
      blocks: nodeChildren(bodyStoryRoot(parsed.part)!),
      lastMarkCovered: true,
    });
    if (!pasted.ok) throw new Error(pasted.reason);
    part = pasted.part;
    // The new cell's generated host tail is not an authored target paragraph.
    const tailId = pasted.effect.caret?.paragraphId;
    if (ast.at(-1)?.type === 'paragraph' && tailId && paragraphTextOf(part, tailId) === '') {
      const removed = removeNode(part, tailId);
      if (!removed.ok) throw new Error('Cannot remove generated new-cell tail');
      part = removed.part;
    }
  }
  return withPart(pkg, part);
}
