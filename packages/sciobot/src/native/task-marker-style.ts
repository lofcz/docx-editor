import {
  applyTreeOp,
  authoredProperties,
  AUTHORABLE_RUN_PROPERTIES,
  contentControlContentNodeOf,
  findNode,
  propertyContainer,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

/** Reserve a visible task gap in Word and native layout without adding a text offset. */
export function withPreparedTaskMarker(
  pkg: OoxmlPackage,
  partName: string,
  paragraphId: string,
  controlId: string
): OoxmlPackage {
  const part = pkg.parts.get(partName)!;
  const control = findNode(part, controlId);
  if (control?.kind !== 'contentControl') return pkg;
  const runs = contentControlContentNodeOf(control)?.children.filter((node) => node.kind === 'run');
  if (runs?.length !== 1) return pkg;
  const run = runs[0]!;
  const properties = authoredProperties(
    propertyContainer(run, 'runProperties', 'rPr'),
    AUTHORABLE_RUN_PROPERTIES
  );
  const additions = [
    { localName: 'strike', attributes: { val: '0' } },
    { localName: 'spacing', attributes: { val: '180' } },
    { localName: 'sz', attributes: { val: '30' } },
  ].filter((property) => !properties.some((existing) => existing.localName === property.localName));
  if (!additions.length) return pkg;
  const changed = applyTreeOp(part, {
    op: 'setRunProperties',
    paragraphId,
    start: 0,
    end: 1,
    targetRunIds: [run.id],
    properties: [...properties, ...additions],
  });
  if (!changed.ok) throw new Error('The task marker formatting was refused.');
  return withPart(pkg, changed.part);
}
