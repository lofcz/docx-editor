import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  createFixedMeasurer,
  paragraphFragmentsOf,
  paragraphFragmentsOfBlocks,
  type ParagraphFragmentRecord,
  type StyleSpanRecord,
} from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  collectRevisionSites,
  contentControlPropertiesOf,
  openHeadlessDocument,
  parentNodeOf,
  paragraphTextOf,
  readOoxmlPackage,
  runPropertyEdits,
  runsCovering,
  storyParagraphs,
  storyRootsOf,
  withPart,
  type HeadlessDocumentView,
  type OoxmlElement,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { checklistTag } from '../registry/adapter-names';
import { runCharacterStyle } from './inline-code';
import { initializedPackageBytes } from './prepared-parts';
import { taskAppearanceContribution } from './task-appearance';
import {
  taskExportStyleKey,
  withTaskExportStyles,
  type TaskExportStyle,
} from './task-export-styles';

export interface TaskStoryParagraph {
  readonly part: OoxmlPart;
  readonly paragraph: OoxmlElement;
  readonly key: string;
  readonly pending: boolean;
}
/** Native story order supplies stable identities across detached serialization. */
export function taskStoryParagraphs(pkg: OoxmlPackage): readonly TaskStoryParagraph[] {
  const result: TaskStoryParagraph[] = [];
  for (const part of pkg.parts.values()) {
    const roots = storyRootsOf(part).filter((story) => story.kind !== 'note');
    if (!roots.length) continue;
    const pending = new Set<string>();
    for (const site of collectRevisionSites(part)) {
      let node: OoxmlElement | null = site.node;
      for (let depth = 0; node && depth < 64; depth++) {
        if (node.kind === 'paragraph' || node.kind === 'table') {
          pending.add(node.id);
          break;
        }
        node = parentNodeOf(part, node.id);
      }
    }
    let index = 0;
    for (const root of roots)
      for (const paragraph of storyParagraphs(root.root)) {
        if (paragraph.kind !== 'paragraph') continue;
        const key = `${part.name}:${index++}`;
        let node: OoxmlElement | null = paragraph,
          hasPending = false;
        for (let depth = 0; node && depth < 64; depth++) {
          if (pending.has(node.id)) {
            hasPending = true;
            break;
          }
          node = parentNodeOf(part, node.id);
        }
        result.push({ part, paragraph, key, pending: hasPending });
      }
  }
  return result;
}
export function isApplicationTask(paragraph: OoxmlElement): boolean {
  return paragraph.children.some((node) => {
    if (node.kind !== 'contentControl') return false;
    const properties = contentControlPropertiesOf(node);
    return properties.tag === checklistTag() && properties.type === 'checkbox';
  });
}
async function appearanceSnapshot(view: HeadlessDocumentView, appearance: boolean) {
  const identities = new Map(
    taskStoryParagraphs(view.currentPackage()).map(({ part, paragraph, key }) => [
      `${part.name}:${paragraph.id}`,
      key,
    ])
  );
  const opened = openDocumentForExport(view, {
    measurer: createFixedMeasurer(),
    displayMode: 'proposed',
    ...(appearance ? { paragraphAppearance: taskAppearanceContribution() } : {}),
  });
  if (!opened.ok) throw new Error(`Cannot lay out task export: ${opened.reason}.`);
  const values = new Map<string, Map<string, StyleSpanRecord>>();
  const add = (partName: string, fragments: readonly ParagraphFragmentRecord[]) => {
    for (const paragraph of fragments) {
      const key = identities.get(`${partName}:${paragraph.paragraphId}`);
      if (!key) continue;
      let spans = values.get(key);
      if (!spans) {
        spans = new Map();
        values.set(key, spans);
      }
      for (const line of paragraph.lines)
        for (const span of line.spans)
          spans.set(
            JSON.stringify([span.range.start, span.range.end, span.text, span.style]),
            span
          );
    }
  };
  try {
    for (const page of (await opened.session.layout()).pages) {
      add(view.currentPackage().mainDocumentPart, paragraphFragmentsOf(page));
      for (const story of [page.header, page.footer])
        if (story) add(story.partName, paragraphFragmentsOfBlocks(story.fragments));
    }
  } finally {
    opened.session.dispose();
  }
  return new Map(
    [...values].map(([key, spans]) => [
      key,
      [...spans.values()].sort(
        (a, b) =>
          a.range.start - b.range.start || a.range.end - b.range.end || a.text.localeCompare(b.text)
      ),
    ])
  );
}
const appearanceKey = (spans: readonly StyleSpanRecord[] | undefined) =>
  JSON.stringify(
    spans?.map((span) => [span.range.start, span.range.end, span.text, span.style]) ?? []
  );
function overlapping(spans: readonly StyleSpanRecord[], start: number, end: number) {
  let lo = 0,
    hi = spans.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (spans[mid]!.range.end <= start) lo = mid + 1;
    else hi = mid;
  }
  const found: StyleSpanRecord[] = [];
  for (let index = lo; index < spans.length && spans[index]!.range.start < end; index++)
    found.push(spans[index]!);
  return found;
}

/**
 * Materialize settled application task appearance in a detached Word download.
 * Pending review paragraphs keep their canonical content and review records unchanged.
 * Their dynamic completion appearance is intentionally not materialized for Word.
 * Native editing and collaboration snapshots must continue using raw editor.save().
 */
export async function materializeTaskAppearanceForWord(document: Uint8Array): Promise<Uint8Array> {
  const opened = openHeadlessDocument(document);
  if (!opened.ok) throw new Error(`Cannot open task export: ${opened.reason}.`);
  // Headless initialization adds paragraph identities. Only the untouched parser result
  // may supply detached writes, or unrelated pending paragraphs gain new attributes.
  const loaded = readOoxmlPackage(document);
  if (!loaded.ok) throw new Error(`Cannot read task export: ${loaded.reason}.`);
  const original = loaded.package;
  const tasks = taskStoryParagraphs(original).filter(
    ({ paragraph, pending }) => !pending && isApplicationTask(paragraph)
  );
  if (!tasks.length) return document;
  const before = await appearanceSnapshot(opened.view, false);
  const desired = await appearanceSnapshot(opened.view, true);
  const plans: {
    partName: string;
    paragraphId: string;
    edit: ReturnType<typeof runPropertyEdits>[number];
    source: TaskExportStyle;
  }[] = [];
  for (const { part, paragraph, key } of tasks) {
    if (appearanceKey(before.get(key)) === appearanceKey(desired.get(key))) continue;
    for (const edit of runPropertyEdits(
      part,
      paragraph.id,
      0,
      paragraphTextOf(part, paragraph.id)?.length ?? 0,
      [],
      'proposed'
    )) {
      const oldSpans = overlapping(before.get(key) ?? [], edit.start, edit.end);
      const newSpans = overlapping(desired.get(key) ?? [], edit.start, edit.end);
      if (!oldSpans.length || !newSpans.length) continue;
      const strikeToggle = oldSpans[0]!.style.strike !== newSpans[0]!.style.strike;
      const color = oldSpans[0]!.style.color !== newSpans[0]!.style.color;
      if (!strikeToggle && !color) continue;
      if (
        oldSpans.some(
          (span) =>
            span.style.strike !== oldSpans[0]!.style.strike ||
            span.style.color !== oldSpans[0]!.style.color
        ) ||
        newSpans.some(
          (span) =>
            span.style.strike !== newSpans[0]!.style.strike ||
            span.style.color !== newSpans[0]!.style.color
        )
      )
        throw new Error('Task export cannot represent conflicting native formatting in one run.');
      const run = runsCovering(part, paragraph.id, edit.start, edit.end, 'proposed').find(
        (node) => !edit.targetRunIds || edit.targetRunIds.includes(node.id)
      );
      if (!run) throw new Error('Task export cannot identify its native source run.');
      plans.push({
        partName: part.name,
        paragraphId: paragraph.id,
        edit,
        source: { base: runCharacterStyle(run), strikeToggle, color },
      });
    }
  }
  if (!plans.length) return document;
  const definitions = withTaskExportStyles(
    original,
    plans.map((plan) => plan.source)
  );
  let next = definitions.pkg;
  for (const plan of plans) {
    const result = applyTreeOp(next.parts.get(plan.partName)!, {
      op: 'setRunProperties',
      paragraphId: plan.paragraphId,
      start: plan.edit.start,
      end: plan.edit.end,
      properties: plan.edit.properties,
      characterStyle: definitions.ids.get(taskExportStyleKey(plan.source))!,
      ...(plan.edit.targetRunIds ? { targetRunIds: plan.edit.targetRunIds } : {}),
    });
    if (!result.ok) throw new Error(`Task export formatting was refused: ${result.reason}.`);
    next = withPart(next, result.part);
  }
  const bytes = initializedPackageBytes(document, original, next);
  const reopened = openHeadlessDocument(bytes);
  if (!reopened.ok) throw new Error(`Cannot reopen task export: ${reopened.reason}.`);
  const actual = await appearanceSnapshot(reopened.view, false);
  for (const { key } of tasks)
    if (appearanceKey(actual.get(key)) !== appearanceKey(desired.get(key)))
      throw new Error('Task export changed native Word formatting; the export was refused.');
  return bytes;
}
