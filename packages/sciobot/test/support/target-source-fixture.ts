import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  escapeXml,
  readOoxmlPackage,
  headlessViewOfStore,
  TreePackageStore,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { blockTag, withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import type { BlockPayload } from '../../src/blocks/payload-schema';
import {
  DOMAIN_BLOCK_TAGS,
  GENERATED_LABEL_TAG,
  type BlockTagDefinition,
} from '../../src/projection/block-tags';
import { MDX_BLOCK, MDX_INLINE, projectMdx } from '../../src/projection/mdx';
import { listDocument, listParagraph, checklistParagraph } from './list-document';
import { formatCell, formatRow, formatTable } from './table-format-fixture';

export const sourceRevision = { documentId: 'target-owners', epoch: 'load-1', revision: 4 };
export const sourceParagraph = (text: string) =>
  `<w:p><w:pPr><keep:meta xmlns:keep="urn:keep" keep:value="paragraph"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
export const sourceControl = (id: string, body: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
const label = (text: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${GENERATED_LABEL_TAG}"/></w:sdtPr><w:sdtContent>${sourceParagraph(text)}</w:sdtContent></w:sdt>`;
export const sourceDefinitions = new Map([
  ...DOMAIN_BLOCK_TAGS,
  ...([
    ['custom_container', { tag: 'custom_box', kind: 'container' as const, attributes: () => [] }],
    ['custom_atomic', { tag: 'custom_atom', kind: 'atomic' as const, attributes: () => [] }],
  ] as const),
]);
const payload = (
  id: string,
  type: string,
  attributes: BlockPayload['attributes'] = {}
): BlockPayload => ({ id, type, attributes, version: 1 });
export async function projectSourcePackage(
  pkg: OoxmlPackage,
  mode: 'proposed' | 'original' | 'all-markup' = 'proposed',
  definitions: ReadonlyMap<string, BlockTagDefinition> = sourceDefinitions
) {
  const main = pkg.parts.get(pkg.mainDocumentPart)!;
  const opened = openDocumentForExport(headlessViewOfStore(new TreePackageStore(pkg, main)));
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const result = projectMdx(
      pkg,
      await opened.session.layoutFor(mode),
      sourceRevision,
      definitions
    );
    if (!result.success) throw new Error(result.error);
    return result.projection;
  } finally {
    opened.session.dispose();
  }
}
export async function targetSourceFixture() {
  const nested = formatTable(formatRow('Nested A', 'Nested B', 'Nested C'));
  const table = formatTable(
    '<w:tr>' +
      formatCell('AB', 3600, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') +
      formatCell('C1', 3600) +
      '</w:tr><w:tr>' +
      formatCell('', 3600, '<w:gridSpan w:val="2"/><w:vMerge/>') +
      formatCell('C2', 3600) +
      '</w:tr><w:tr><w:tc><w:tcPr/>' +
      sourceParagraph('Outer cell') +
      nested +
      '</w:tc>' +
      formatCell('B3') +
      formatCell('C3', 3600) +
      '</w:tr>'
  );
  const body =
    sourceParagraph('Before') +
    '<keep:block xmlns:keep="urn:keep" keep:value="barrier"><keep:data>Opaque</keep:data></keep:block>' +
    sourceControl(
      'phase',
      label('Phase label') +
        sourceControl(
          'activity',
          label('Activity label') +
            sourceParagraph('Question') +
            table +
            sourceControl('opaque', sourceParagraph('Atomic fallback')) +
            sourceParagraph('Activity tail')
        )
    ) +
    '<w:p><w:r><w:t xml:space="preserve">Inline before </w:t></w:r>' +
    sourceControl('inline', '<w:r><w:t>Inline fallback</w:t></w:r>') +
    '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>' +
    listParagraph('First', '1') +
    listParagraph('Second', '1') +
    checklistParagraph(false).replace('>Task<', '><') +
    sourceControl('custom_container', sourceParagraph('')) +
    sourceControl('custom_atomic', sourceParagraph('Custom fallback')) +
    '<w:sdt><w:sdtPr><w:alias w:val="Ordinary Word control"/></w:sdtPr><w:sdtContent>' +
    sourceParagraph('Word form text') +
    '</w:sdtContent></w:sdt>' +
    sourceParagraph('Tail') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
  const loaded = readOoxmlPackage(listDocument(body));
  if (!loaded.ok) throw new Error(loaded.reason);
  let pkg = loaded.package;
  const values: BlockPayload[] = [
    payload('phase', 'lesson_phase', { phaseType: 'evocation' }),
    payload('activity', 'lesson_activity', { name: 'Warmup', duration: 10 }),
    payload('custom_container', 'custom_container'),
    payload('custom_atomic', 'custom_atomic'),
    {
      ...payload('opaque', MDX_BLOCK),
      mdast: {
        type: 'mdxJsxFlowElement',
        name: 'UnknownPanel',
        attributes: [{ type: 'mdxJsxExpressionAttribute', value: '...danger()' }],
        children: [{ type: 'mdxFlowExpression', value: 'danger()' }],
        data: { future: { type: 'payload', position: { keep: true } } },
      },
    },
    {
      ...payload('inline', MDX_INLINE),
      mdast: {
        type: 'mdxJsxTextElement',
        name: 'Badge',
        attributes: [{ type: 'mdxJsxAttribute', name: 'label', value: 'Keep' }],
        children: [],
        future: { keep: true },
      },
    },
  ];
  for (const value of values) {
    const stored = withPayloadRecord(pkg, value, '');
    if (!stored.ok) throw new Error(stored.reason);
    pkg = stored.pkg;
  }
  const projection = await projectSourcePackage(pkg);
  return { revision: sourceRevision, pkg, projection };
}
