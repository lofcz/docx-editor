import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { docxFixture } from './document';
import { propertyContext } from './property-execution';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { readWorkingView } from '../../src/review/working-view';
import { recomposeOwnedProposals } from '../../src/review/recompose';
import { setManualSuggestionMode } from '../../src/review/manual-tracking';

export const manualContext = {
  ...propertyContext,
  author: 'ScioBot AI',
  date: '2026-10-24T12:00:00.000Z',
};
export type ManualWave = Extract<
  Awaited<ReturnType<typeof recomposeOwnedProposals>>,
  { success: true }
>;
export function manualTrackingDocument() {
  return docxFixture(
    ['Alpha end', 'Manual text', 'Delete this', 'Bold text', 'Tail']
      .map((text, index) => {
        const id = (index + 1).toString(16).padStart(8, '0');
        return `<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${id}" w14:textId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
      })
      .join('')
  );
}
export const manualToken = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, manualContext.documentId, manualContext.epoch)!;
export function selectManual(
  editor: DocxEditorInstance,
  index: number,
  start: number,
  end = start
) {
  const s = editor.surface!,
    paragraphId = s.session.paragraphIds()[index]!;
  s.setSelection({ anchor: { paragraphId, offset: start }, head: { paragraphId, offset: end } });
}
export async function manualWorking(editor: DocxEditorInstance) {
  const result = await readWorkingView(editor.surface!.session, manualToken(editor));
  if (!result.success) throw new Error(result.error);
  return result.view.projection.text;
}
export async function manualWave(
  editor: DocxEditorInstance,
  value: string,
  previous?: ManualWave,
  author = manualContext.author
) {
  const target = (await manualWorking(editor)).replace(
    /Alpha (?:end|ONE end|TWO end)/,
    `Alpha ${value} end`
  );
  const result = await recomposeOwnedProposals(
    editor,
    {
      revision: manualToken(editor),
      targetMarkdown: target,
      ownership: previous?.ownership ?? [],
      contentProofs: previous?.contentProofs ?? [],
    },
    { ...manualContext, author }
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result;
}
export function manualEdits(editor: DocxEditorInstance) {
  const enabled = setManualSuggestionMode(editor, true);
  if (!enabled.ok) throw new Error(enabled.reason);
  const s = editor.surface!;
  selectManual(editor, 1, 7);
  s.insertPlainText('new ');
  s.flushPendingInput();
  selectManual(editor, 2, 0, 7);
  s.deleteForward();
  selectManual(editor, 3, 0, 4);
  const bold = editor.exec({ type: 'toggleMark', mark: 'bold' });
  if (!bold.ok) throw new Error(bold.reason);
}
const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find(
        (a) =>
          a.namespaceUri === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' &&
          a.localName === name
      )?.value;
export function manualRevisionFingerprints(
  editor: DocxEditorInstance,
  author = editor.getConfiguredAuthor()
) {
  return collectRevisionSites(editor.surface!.session.part())
    .filter((site) => attr(site.node, 'author') === author)
    .map((site) => ({
      nodeId: site.node.id,
      id: attr(site.node, 'id'),
      kind: site.node.kind,
      fingerprint: canonicalOoxmlFingerprint(site.node),
    }));
}

/** Include proposed formatting and all surrounding native content, beyond the old-value record. */
export function manualParagraphFingerprints(editor: DocxEditorInstance) {
  const s = editor.surface!;
  return s.session
    .paragraphIds()
    .slice(1, 4)
    .map((id) => canonicalOoxmlFingerprint(findNode(s.session.part(), id)!));
}
