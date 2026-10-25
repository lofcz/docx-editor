import {
  buildNumberingIndex,
  buildStyleCascadeTable,
  cascadeParagraphFormatting,
  createListCounterState,
  readNumPr,
  resolveNumberingLevel,
  withNumberingStyleLinks,
} from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  propertyContainer,
  resolveRelationship,
  storyParagraphs,
  stylesPartOf,
} from '@docx-editor.dev/core/store';
import type { ListMembershipPatch, ListPlanSource } from './lists';

/** Read parent counter values with the native counter engine, including preceding edits. */
export function listCounterAncestry(
  source: ListPlanSource,
  patches: readonly ListMembershipPatch[] = []
) {
  const pkg = source.pkg,
    part = pkg.parts.get(source.partName)!;
  const relationship = pkg.relationships
    .get(pkg.mainDocumentPart)
    ?.find(
      (record) =>
        record.type ===
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
    );
  const resolved = relationship && resolveRelationship(relationship);
  const numberingPart =
    (resolved?.mode === 'Internal' && resolved.target.ok
      ? pkg.parts.get(resolved.target.partName)
      : undefined) ?? pkg.parts.get('/word/numbering.xml');
  const styles = buildStyleCascadeTable(stylesPartOf(pkg)?.root ?? null);
  const index = withNumberingStyleLinks(buildNumberingIndex(numberingPart?.root), styles);
  const counters = createListCounterState(index);
  const targets = new Map(patches.map((patch) => [patch.paragraphId, patch.target]));
  const ancestry = new Map<string, readonly number[]>();
  for (const paragraph of storyParagraphs(bodyStoryRoot(part) ?? part.root)) {
    const target = targets.get(paragraph.id);
    if (target?.kind === 'none' || target?.kind === 'checklist') continue;
    if (target && target.instance.kind !== 'existing')
      throw new Error('Counter ancestry needs allocated native list instances.');
    const member =
      target && target.instance.kind === 'existing'
        ? { numId: target.instance.numId, ilvl: target.level }
        : readNumPr(
            cascadeParagraphFormatting(
              styles,
              propertyContainer(paragraph, 'paragraphProperties', 'pPr')
            ).paragraphPropertyNodes
          );
    if (!member || member.numId === '0') continue;
    const advanced = counters.advance(member.numId, member.ilvl);
    if (!advanced) continue;
    const parents = Array.from({ length: member.ilvl }, (_, level) => {
      const value = advanced.counters[level];
      const native = resolveNumberingLevel(index, member.numId, level);
      return value ? value : (native?.startOverride ?? native?.level.start ?? 1);
    });
    ancestry.set(paragraph.id, parents);
  }
  return ancestry;
}
