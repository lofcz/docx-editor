import {
  collectRevisionSites,
  bodyStoryRoot,
  contentControlPropertiesOf,
  parentNodeOf,
  storyParagraphs,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createListSourceReader } from '../diff/lists';
import { checklistTag } from '../registry/adapter-names';
import { propertyControlRefusal } from '../review/property-fallback';
import { applyTaskCompletionStyle } from './task-completion-style';
import { withPreparedTaskMarker } from './task-marker-style';

/** Restore owned task appearance before room creation without changing imported Word checkboxes. */
export function withPreparedTaskCompletion(pkg: OoxmlPackage): OoxmlPackage {
  const part = pkg.parts.get(pkg.mainDocumentPart);
  if (!part) return pkg;
  const body = bodyStoryRoot(part);
  if (!body) return pkg;
  const pending = new Set<string>();
  for (const site of collectRevisionSites(part)) {
    let node = site.parent;
    for (let depth = 0; node && depth < 64; depth++) {
      if (node.kind === 'paragraph') {
        pending.add(node.id);
        break;
      }
      node = parentNodeOf(part, node.id);
    }
  }
  const reader = createListSourceReader(pkg, part);
  let next = pkg;
  for (const paragraph of storyParagraphs(body)) {
    if (paragraph.kind !== 'paragraph') continue;
    if (pending.has(paragraph.id)) continue;
    const owned = paragraph.children.find(
      (node) =>
        node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === checklistTag()
    );
    if (!owned || propertyControlRefusal(part, owned.id)) continue;
    const source = reader.read(paragraph.id);
    if (!source.success || !source.state.checkbox) continue;
    next = withPreparedTaskMarker(next, part.name, paragraph.id, owned.id);
    next = applyTaskCompletionStyle(next, part.name, paragraph.id, source.state.checkbox.checked);
  }
  return next;
}
