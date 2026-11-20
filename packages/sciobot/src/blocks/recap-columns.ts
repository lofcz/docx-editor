import {
  bodyStoryRoot,
  contentControlContentNodeOf,
  createNodeIdAllocator,
  findNode,
  readOoxmlPart,
  replaceChildren,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from './block-index';
import { columnTableXml, readPresentationColumns, recapColumnPair } from './column-geometry';

/** Prepare older lesson overviews before a shared room opens. Keep both original controls intact. */
export function withRecapColumns(pkg: OoxmlPackage): OoxmlPackage {
  const index = buildBlockIndex(pkg);
  const byControl = new Map([...index.byId.values()].map((block) => [block.controlId, block]));
  let result = pkg;
  for (const recap of index.byId.values()) {
    if (recap.payload.type !== 'preparation_recap' || readPresentationColumns(result, recap, index))
      continue;
    const part = result.parts.get(recap.partName);
    const control = part && findNode(part, recap.controlId);
    const content = control && contentControlContentNodeOf(control);
    if (!part || !content) continue;
    const at = content.children.findIndex((node, i) =>
      recapColumnPair(
        byControl.get(node.id)?.payload.type,
        byControl.get(content.children[i + 1]?.id ?? '')?.payload.type
      )
    );
    if (at < 0) continue;
    const pair = content.children.slice(at, at + 2);
    const materialsColumn = byControl.get(pair[0]!.id)?.payload.type === 'recap_materials' ? 0 : 1;
    const skeleton = columnTableXml([4320, 4320], () => '<w:p/>', materialsColumn);
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${skeleton}</w:body></w:document>`,
      {
        name: '/recap-columns.xml',
        contentType: part.contentType,
      }
    );
    const body = parsed.ok && bodyStoryRoot(parsed.part);
    const table = body && body.kind !== 'textValue' && body.children[0];
    if (!table || table.kind !== 'table') throw new Error('Cannot create the overview columns.');
    const next = createNodeIdAllocator(part);
    let column = 0;
    const fill = (node: OoxmlNode): OoxmlNode => {
      if (node.kind === 'paragraph') return pair[column++]!;
      if (node.kind === 'textValue') return { ...node, id: next() };
      return { ...node, id: next(), children: node.children.map(fill) } as OoxmlNode;
    };
    const updated = replaceChildren(part, content.id, [
      ...content.children.slice(0, at),
      fill(table),
      ...content.children.slice(at + pair.length),
    ]);
    if (!updated.ok) throw new Error('Cannot prepare the overview columns.');
    result = withPart(result, updated.part);
  }
  return result;
}
