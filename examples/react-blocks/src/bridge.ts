import {
  createBrowserAutomationHost,
  type DocxEditorInstance,
  type EditorModule,
} from '@docx-editor.dev/core/editor';
import {
  collectReviewItems,
  revisionItemsOf,
  paragraphTextOf,
  findNode,
} from '@docx-editor.dev/core/store';
import type {
  AutomationHandle,
  AutomationBatchResponse,
  AutomationOperation,
} from '@docx-editor.dev/core/automation';
import { attribute, elementChild, type BlockData, type DocumentBlock } from './document';

export const REVIEW: EditorModule = {
  id: 'scio-review',
  review: {
    displayModes: ['all-markup', 'proposed', 'original'],
    collectReviewItems,
    revisionItemsOfParagraph: (part, id) =>
      revisionItemsOf(part).filter((item) =>
        item.ranges.every((range) => range.start.paragraphId === id && range.end.paragraphId === id)
      ),
  },
};
const value = (response: AutomationBatchResponse, index = 0) => {
  const result = response.results[index];
  if (!response.ok || result?.status !== 'ok') throw new Error(JSON.stringify(response));
  return result.value;
};

export function createBlockBridge(editor: DocxEditorInstance) {
  const host = createBrowserAutomationHost(editor);
  // Each operation is planned from live synchronous reads. The browser host
  // flushes buffered native input before it executes the batch.
  const execute = (operations: AutomationOperation[]) => host.execute({ operations });
  const doc = value(execute([{ op: 'getDocument' }]));
  if (doc.kind !== 'handle') throw new Error('No document');
  const bodyResult = value(execute([{ op: 'getBody', document: doc.handle }]));
  if (bodyResult.kind !== 'handle') throw new Error('No body');
  const body = bodyResult.handle;
  function paragraph(wordId: string): AutomationHandle {
    const found = value(execute([{ op: 'getParagraphs', body }]));
    if (found.kind !== 'handles') throw new Error('No paragraphs');
    const ids = execute(found.handles.map((paragraph) => ({ op: 'getParagraphId', paragraph })));
    const index = ids.results.findIndex((_, i) => {
      const item = value(ids, i);
      return item.kind === 'text' && item.text === wordId;
    });
    if (index < 0) throw new Error(`Missing paragraph ${wordId}`);
    return found.handles[index]!;
  }
  function update(block: DocumentBlock, data: BlockData) {
    const controls = value(
      execute([{ op: 'getContentControlsByTag', scope: { body }, tag: block.tag }])
    );
    if (controls.kind !== 'handles' || controls.handles.length !== 1)
      throw new Error('Ambiguous block');
    const header = paragraph(block.headerWordId);
    const oldText = paragraphTextOf(editor.surface!.session.part(), block.headerId) ?? '';
    const nextText =
      block.type === 'activity'
        ? `${data.name} — ${data.duration} minutes`
        : block.type === 'quiz'
          ? `${data.name}. Prediction: ${data.answer || 'not chosen'}.`
          : data.name;
    value(
      execute([
        {
          op: 'setContentControlProperties',
          contentControl: controls.handles[0]!,
          title: JSON.stringify(data),
        },
        {
          op: 'replaceSpan',
          span: {
            start: { paragraph: header, offset: 0 },
            end: { paragraph: header, offset: oldText.length },
          },
          text: nextText,
        },
      ])
    );
  }
  function resize(block: DocumentBlock, heightPx: number, scale: number) {
    const heightPt = Math.ceil(heightPx / scale + 8);
    const header = paragraph(block.headerWordId);
    const node = findNode(editor.surface!.session.part(), block.headerId);
    const props = node && node.kind !== 'textValue' ? elementChild(node, 'pPr') : undefined;
    const spacing = elementChild(props, 'spacing');
    if (Math.abs(Number(attribute(spacing, 'line')) / 20 - heightPt) < 2) return;
    const result = execute([
      {
        op: 'setParagraphFormat',
        paragraph: { paragraph: header },
        format: { lineSpacing: heightPt },
      },
    ]);
    if (!result.ok) throw new Error(JSON.stringify(result));
  }
  function propose() {
    const spans = value(execute([{ op: 'search', scope: { body }, text: 'Plan an experiment.' }]));
    if (spans.kind !== 'spans' || spans.spans.length !== 1)
      throw new Error('Expected one editable activity paragraph');
    value(
      execute([
        {
          op: 'proposeReplacement',
          author: 'ScioBot',
          span: spans.spans[0]!,
          text: 'Plan a controlled experiment.',
        },
      ])
    );
  }
  function pushBlocks() {
    const matches = value(
      execute([
        { op: 'search', scope: { body }, text: 'Ordinary editable text before the custom blocks.' },
      ])
    );
    if (matches.kind !== 'spans' || matches.spans.length !== 1)
      throw new Error('Expected one prelude');
    value(
      execute([
        {
          op: 'setParagraphFormat',
          paragraph: { paragraph: matches.spans[0]!.start.paragraph },
          format: { spaceAfter: 300 },
        },
      ])
    );
  }
  return { update, resize, propose, pushBlocks, destroy: () => host.dispose() };
}
export type BlockBridge = ReturnType<typeof createBlockBridge>;
