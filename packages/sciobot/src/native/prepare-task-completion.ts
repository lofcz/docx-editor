import {
  contentControlPropertiesOf,
  type OoxmlPackage,
  applyTreeOp,
  paragraphTextOf,
  runPropertyEdits,
  runsCovering,
  withPart,
} from '@docx-editor.dev/core/store';
import { createListSourceReader } from '../diff/lists';
import { checklistTag } from '../registry/adapter-names';
import { propertyControlRefusal } from '../review/property-fallback';
import { restoreLegacyTaskCompletionStyle } from './task-completion-style';
import { withPreparedTaskMarker } from './task-marker-style';
import { taskExportStyleSource } from './task-export-styles';
import { runCharacterStyle } from './inline-code';
import { taskStoryParagraphs } from './task-export';

/** Restore authored task styles after Word materialization; live appearance follows checkbox state. */
export function withPreparedTaskCompletion(pkg: OoxmlPackage): OoxmlPackage {
  const readers = new Map<string, ReturnType<typeof createListSourceReader>>();
  let next = pkg;
  for (const { part, paragraph, pending } of taskStoryParagraphs(pkg)) {
    if (pending) continue;
    const owned = paragraph.children.find(
      (node) =>
        node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === checklistTag()
    );
    if (!owned || propertyControlRefusal(part, owned.id)) continue;
    let reader = readers.get(part.name);
    if (!reader) {
      reader = createListSourceReader(pkg, part);
      readers.set(part.name, reader);
    }
    const source = reader.read(paragraph.id);
    if (!source.success || !source.state.checkbox) continue;
    next = withPreparedTaskMarker(next, part.name, paragraph.id, owned.id);
    next = restoreLegacyTaskCompletionStyle(next, part.name, paragraph.id);
    const current = next.parts.get(part.name)!;
    for (const edit of runPropertyEdits(
      current,
      paragraph.id,
      0,
      paragraphTextOf(current, paragraph.id)?.length ?? 0,
      [],
      'proposed'
    )) {
      const run = runsCovering(current, paragraph.id, edit.start, edit.end, 'proposed').find(
        (node) => !edit.targetRunIds || edit.targetRunIds.includes(node.id)
      );
      const source = run && taskExportStyleSource(next, runCharacterStyle(run));
      if (!source) continue;
      const restored = applyTreeOp(next.parts.get(part.name)!, {
        op: 'setRunProperties',
        paragraphId: paragraph.id,
        start: edit.start,
        end: edit.end,
        properties: edit.properties,
        characterStyle: source.base ?? null,
        ...(edit.targetRunIds ? { targetRunIds: edit.targetRunIds } : {}),
      });
      if (!restored.ok) throw new Error('Task export style restoration was refused.');
      next = withPart(next, restored.part);
    }
  }
  return next;
}
