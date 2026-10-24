import {
  readOoxmlPackage,
  replaceNode,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from './document';
import { checklistParagraph, listPPr } from './list-document';

/** Static Word markup derived from the frozen source action, independent of its writer. */
export function checklistIndentOracle(
  source: OoxmlPart,
  sourceId: string,
  prior: boolean,
  accepted: boolean,
  numId: string,
  level: number,
  text: 'Alpha beta' | 'Alpha peer beta'
): OoxmlPart {
  const markup = accepted
    ? '<w:p><w:pPr>' +
      listPPr(numId, level) +
      (prior ? '<w:customKeep w:val="yes"/>' : '') +
      '<scioList:state xmlns:scioList="urn:sciobot:docx:lists:1" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="scioList" mc:PreserveElements="scioList:state" checked="false"/>' +
      `</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    : prior
      ? checklistParagraph(false).replace('Task', text)
      : `<w:p><w:pPr/><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const loaded = readOoxmlPackage(
    docxFixture(
      markup.replace(
        '<w:p>',
        '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
      )
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const part = loaded.package.parts.get(loaded.package.mainDocumentPart)!;
  const find = (root: OoxmlNode, match: (node: OoxmlNode) => boolean): OoxmlNode | undefined => {
    if (match(root)) return root;
    if (root.kind !== 'textValue')
      for (const child of root.children) {
        const found = find(child, match);
        if (found) return found;
      }
  };
  const original = find(source.root, (node) => node.id === sourceId),
    expected = find(part.root, (node) => node.kind === 'paragraph');
  if (original?.kind !== 'paragraph' || expected?.kind !== 'paragraph')
    throw new Error('Missing oracle paragraph.');
  const result = replaceNode(part, expected.id, { ...expected, attributes: original.attributes });
  if (!result.ok) throw new Error('Invalid oracle paragraph identity.');
  return result.part;
}
