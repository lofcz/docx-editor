import type { SemanticLayout } from '../layout/semantic-records.ts';
import { paragraphFragmentsOnPage } from '../layout/story-fragments.ts';

const atoms = new WeakMap<SemanticLayout, ReadonlySet<string>>();

/** A host-admitted box is one native input unit. Canonical locks remain independent. */
export function isMeasuredInlineControl(layout: SemanticLayout, controlId: string): boolean {
  let found = atoms.get(layout);
  if (!found) {
    const ids = new Set<string>();
    for (const page of layout.pages)
      for (const paragraph of paragraphFragmentsOnPage(page))
        for (const line of paragraph.lines)
          for (const span of line.spans)
            if (span.inlineWidget) ids.add(span.inlineWidget.controlId);
    found = ids;
    atoms.set(layout, found);
  }
  return found.has(controlId);
}
