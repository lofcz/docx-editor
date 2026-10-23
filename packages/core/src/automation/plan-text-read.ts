import type { AutomationOperation } from './operations.ts';
import type { AutomationHandleTable } from './handles.ts';
import { PARAGRAPH_MARK, type AutomationPackageReads } from './reads.ts';
import type { PlannedOperation } from './plan-types.ts';
import type { AutomationErrorCode, AutomationValue } from './protocol.ts';
import { resolveParagraphHandle, resolveSpanRef, spanText, storyOfHandle } from './spans.ts';
const refuse = (code: AutomationErrorCode, message: string, detail?: string): PlannedOperation => ({
  ok: false,
  error: { code, message, ...(detail ? { detail } : {}) },
});
const query = (value: AutomationValue): PlannedOperation => ({ ok: true, kind: 'query', value });

/** Read a logical paragraph through its canonical fragment range. */
export function planTextRead(
  operation: Extract<AutomationOperation, { op: 'getText' }>,
  handles: AutomationHandleTable,
  packageReads: AutomationPackageReads
): PlannedOperation {
  if (
    !(
      operation.projection === undefined ||
      ['allMarkup', 'original', 'model'].includes(operation.projection)
    )
  )
    return refuse('unsupported-content', 'unknown text projection', 'projection');
  const body = handles.resolve(operation.target, 'body');
  if (body) {
    const story = storyOfHandle(operation.target, 'body', handles, packageReads);
    if (!story.ok) return refuse(story.code, 'that handle does not name a body', story.detail);
    return query({
      kind: 'text',
      text: story.value.text(operation.projection ?? 'allMarkup'),
    });
  }
  const paragraph = resolveParagraphHandle(operation.target, handles, packageReads);
  if (!paragraph.ok)
    return refuse(
      paragraph.code,
      'that handle does not name a body or a paragraph',
      paragraph.detail
    );
  const story = packageReads.story(paragraph.value.story);
  const range = resolveSpanRef({ paragraph: operation.target }, handles, packageReads);
  if (!range.ok) return refuse(range.code, 'that paragraph has no text range', range.detail);
  return query({
    kind: 'text',
    text: story
      ? spanText(range.value, story, PARAGRAPH_MARK, operation.projection ?? 'allMarkup')
      : '',
  });
}
