import {
  bodyStoryRoot,
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  contentControlsIn,
  createNodeIdAllocator,
  escapeXml,
  findNode,
  parentNodeOf,
  readOoxmlPart,
  usedParaIds,
  w14RootPrefix,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { columnTableXml } from './column-geometry';
import { withFreshIds } from './fresh-ids';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import type { IndexedBlock } from './block-index';
import type { BlockPayload } from './payload-schema';

/** New column scaffolding contains no editable paragraph outside its control. */
export function columnScaffold(part: OoxmlPart, widths: readonly number[], ids: readonly string[]) {
  const xml = columnTableXml(
    widths,
    (index) =>
      `<w:sdt><w:sdtPr><w:tag w:val="${escapeXml(blockTag(ids[index]!))}"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`
  );
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${xml}</w:body></w:document>`,
    { name: part.name, contentType: part.contentType }
  );
  const root = parsed.ok ? bodyStoryRoot(parsed.part) : null;
  if (!root || root.kind === 'textValue') throw Error('The column scaffold is invalid.');
  const table = withFreshIds(root.children[0]!, createNodeIdAllocator(part, 'new'), {
    prefix: w14RootPrefix(part.root),
    used: new Set(usedParaIds(part.root)),
    seed: ids.join(':'),
  });
  if (table.kind !== 'table') throw Error('The column scaffold has no table.');
  const grid = table.children.find((node) => node.localName === 'tblGrid');
  const row = table.children.find((node) => node.kind === 'tableRow');
  if (!grid || !row || row.kind !== 'tableRow') throw Error('The column scaffold has no grid.');
  const cells = row.children.filter((node) => node.kind === 'tableCell');
  return { table, grid, row, cells };
}

/** Structural changes preserve locks and bindings throughout the affected group. */
export function columnStructureRefusal(part: OoxmlPart, rootId: string): string | null {
  const root = findNode(part, rootId);
  if (!root) return 'The selected column content is missing.';
  const controls: OoxmlNode[] = contentControlsIn(root).map((entry) => entry.node);
  for (let node: OoxmlNode | null = root; node; node = parentNodeOf(part, node.id))
    if (node.kind === 'contentControl') controls.push(node);
  for (const control of controls) {
    const props = contentControlPropertiesOf(control);
    if (props.dataBinding || (props.lock !== undefined && props.lock !== 'unlocked'))
      return 'The column structure contains protected content.';
  }
  return null;
}

export function columnContent(part: OoxmlPart, controlId: string) {
  const node = findNode(part, controlId);
  const content = node && contentControlContentNodeOf(node);
  if (!content) throw Error('The column content is missing.');
  return content;
}

/** Update only width metadata. Retain every other payload field and readable label. */
export function writeColumnWidths(
  pkg: OoxmlPackage,
  blocks: readonly IndexedBlock[],
  widths: readonly string[]
): OoxmlPackage {
  let next = pkg;
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') throw Error('The column payload store is unavailable.');
  for (const [index, block] of blocks.entries()) {
    const width = widths[index]!;
    const payload: BlockPayload = {
      ...block.payload,
      attributes: { ...block.payload.attributes, width },
    };
    const entry = store.entries.get(block.id);
    if (!entry) throw Error('The column payload is missing.');
    const saved = withPayloadRecord(next, payload, entry.label);
    if (!saved.ok) throw Error(saved.reason);
    next = saved.pkg;
  }
  return next;
}
