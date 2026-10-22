import {
  collectRevisionSites,
  directParagraphProperties,
  ensureHyperlinkRelationship,
  findNode,
  hyperlinkTargetOf,
  relationshipTargetIn,
  propertyContainer,
  runsCovering,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit, PlannedValue } from '../diff/plan';
import { withInlineCodeStyle } from '../native/inline-code';
import type { ReviewGroupReceipt } from './groups';
import {
  nativePropertyAddress,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import { trackedRunFormatOps } from './tracked-run-format';
import type { StagedTextCoordinates } from './tracked-stage-text';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type Format = Extract<CanonicalEdit, { kind: 'format' }>;
const paragraphProperties = new Set([
  'pStyle',
  'jc',
  'spacing',
  'ind',
  'keepNext',
  'keepLines',
  'widowControl',
  'pageBreakBefore',
  'contextualSpacing',
  'bidi',
  'shd',
]);
function attributes(value: PlannedValue): Record<string, string> | null {
  if (
    !value.present ||
    value.value === null ||
    typeof value.value !== 'object' ||
    Array.isArray(value.value)
  )
    return null;
  return Object.values(value.value).every((item) => typeof item === 'string')
    ? (value.value as Record<string, string>)
    : null;
}
function sameAttributes(
  current: Readonly<Record<string, string>> | undefined,
  value: PlannedValue
) {
  if (!value.present) return current === undefined;
  const before = attributes(value);
  return (
    !!current &&
    !!before &&
    Object.keys(current).length === Object.keys(before).length &&
    Object.keys(current).every((key) => Object.hasOwn(before, key) && current[key] === before[key])
  );
}

/** Lower format deltas only after dependent native text allocation. */
export function stageTrackedFormat(
  writer: TrackedStageWriter,
  edit: Format,
  revision: TrackedStageAttribution,
  coordinates?: StagedTextCoordinates,
  unselectedSites: ReadonlySet<string> = new Set()
): readonly ReviewGroupReceipt[] {
  let part = currentStagePart(writer, edit.partName);
  const target = edit.target;
  const id = target.kind === 'hyperlink' ? target.hyperlinkId : target.paragraphId;
  const protectedContent = propertyControlRefusal(part, id);
  if (protectedContent) refuseWrite(protectedContent, 'unsupported-operation');
  const originalSites = new Set(collectRevisionSites(part).map((site) => site.node.id));
  const receipts = (): readonly ReviewGroupReceipt[] => {
    const fresh = collectRevisionSites(currentStagePart(writer, edit.partName)).filter(
      (site) => !originalSites.has(site.node.id)
    );
    if (!fresh.length) return [];
    if (fresh.some((site) => site.refused || !nativePropertyAddress(site.node)))
      refuseWrite('Missing native format review receipt.', 'apply-failed');
    // Native property operations share one transaction-wide address. Separate unrelated
    // paragraph intentions before their exact receipts enter persistent group storage.
    writer.apply(edit.partName, {
      op: 'setRevisionAttribution',
      revision,
      sites: fresh.map((site) => ({
        nodeId: site.node.id,
        renumber: `format:${site.node.localName}`,
      })),
    });
    const addresses = new Map<string, RevisionAddress>();
    for (const site of collectRevisionSites(currentStagePart(writer, edit.partName))) {
      if (originalSites.has(site.node.id)) continue;
      const address = nativePropertyAddress(site.node)!;
      addresses.set(propertyAddressKey(edit.partName, address), address);
    }
    return [{ partName: edit.partName, addresses: [...addresses.values()] }];
  };
  if (target.kind === 'run') {
    const slices = coordinates?.slices(target.start, target.end) ?? [target];
    for (const slice of slices)
      for (const run of runsCovering(
        part,
        target.paragraphId,
        slice.start,
        slice.end,
        'all-markup'
      )) {
        const properties = propertyContainer(run, 'runProperties', 'rPr');
        if (
          properties?.kind !== 'textValue' &&
          properties?.children.some((node) => unselectedSites.has(node.id))
        )
          refuseWrite(
            'Replacement formatting overlaps an unselected property revision.',
            'ambiguous-target'
          );
      }
    // A code delta can create the shared style definition. This package edit stays in this unit.
    const codeStyles = edit.changes
      .flatMap((change) =>
        change.name === 'rStyle' && change.after.present ? [attributes(change.after)?.val] : []
      )
      .filter((value): value is string => value !== undefined);
    if (codeStyles.length)
      writer.edit((pkg) => {
        const styled = withInlineCodeStyle(pkg, codeStyles);
        if (!styled)
          refuseWrite('The document cannot gain the inline code style.', 'unsupported-operation');
        return styled;
      });
    part = currentStagePart(writer, edit.partName);
    const prepared = trackedRunFormatOps(
      part,
      slices.map((slice) => ({
        ...edit,
        target: { ...target, start: slice.start, end: slice.end },
      })),
      revision
    );
    if (!prepared.success) refuseWrite(prepared.error, prepared.code);
    for (const op of prepared.ops) writer.apply(part.name, op);
    return receipts();
  }
  if (target.kind === 'paragraph') {
    if (findNode(part, target.paragraphId)?.kind !== 'paragraph')
      refuseWrite('The native paragraph no longer exists.', 'stale-revision');
    const paragraph = findNode(part, target.paragraphId)!;
    const properties = propertyContainer(paragraph, 'paragraphProperties', 'pPr');
    if (
      properties?.kind !== 'textValue' &&
      properties?.children.some((node) => unselectedSites.has(node.id))
    )
      refuseWrite(
        'Replacement formatting overlaps an unselected property revision.',
        'ambiguous-target'
      );
    const bag = new Map(
      directParagraphProperties(part, target.paragraphId).map((entry) => [entry.localName, entry])
    );
    for (const change of edit.changes) {
      if (
        !paragraphProperties.has(change.name) ||
        (change.before.present && !attributes(change.before)) ||
        (change.after.present && !attributes(change.after))
      )
        refuseWrite('Unsupported native paragraph property delta.', 'unsupported-operation');
      const old = bag.get(change.name);
      if (!sameAttributes(old ? (old.attributes ?? {}) : undefined, change.before))
        refuseWrite(
          'Planned paragraph properties no longer match their before-values.',
          'stale-revision'
        );
      if (change.after.present)
        bag.set(change.name, { localName: change.name, attributes: attributes(change.after)! });
      else bag.delete(change.name);
    }
    writer.apply(part.name, {
      op: 'setParagraphProperties',
      paragraphId: target.paragraphId,
      properties: [...bag.values()],
      revision,
    });
    return receipts();
  }
  const link = findNode(part, target.hyperlinkId);
  if (link?.kind !== 'hyperlink')
    refuseWrite('The planned hyperlink no longer exists.', 'stale-revision');
  const current = hyperlinkTargetOf(link, (id) =>
    relationshipTargetIn(writer.currentPackage(), part.name, id)
  );
  let next: { relationshipId?: string; anchor?: string; tooltip?: string } = {
    ...(current.relationshipId !== undefined ? { relationshipId: current.relationshipId } : {}),
    ...(current.anchor !== undefined ? { anchor: current.anchor } : {}),
  };
  let url: string | undefined;
  for (const change of edit.changes) {
    if (change.name === 'target') {
      const before = change.before.present
        ? (change.before.value as Record<string, unknown>)
        : null;
      const after = change.after.present ? (change.after.value as Record<string, unknown>) : null;
      if (!before || before.kind !== current.kind || before.authored !== current.authored)
        refuseWrite('The planned hyperlink target no longer matches.', 'stale-revision');
      if (
        !after ||
        typeof after.authored !== 'string' ||
        !['external', 'internal'].includes(String(after.kind))
      )
        refuseWrite(
          'A hyperlink target needs a native external URL or bookmark.',
          'unsupported-operation'
        );
      // A destination change retains any tooltip delta that precedes it.
      const tooltip = next.tooltip === undefined ? {} : { tooltip: next.tooltip };
      if (after.kind === 'external') {
        url = after.authored;
        next = tooltip;
      } else next = { ...tooltip, anchor: after.authored };
    } else if (change.name === 'tooltip') {
      if (
        change.before.present
          ? change.before.value !== current.tooltip
          : current.tooltip !== undefined
      )
        refuseWrite('The planned hyperlink tooltip no longer matches.', 'stale-revision');
      if (!change.after.present || typeof change.after.value !== 'string')
        refuseWrite(
          'Removing a hyperlink tooltip needs native link structure planning.',
          'unsupported-operation'
        );
      next.tooltip = change.after.value;
    } else refuseWrite('Unknown native hyperlink change.', 'unsupported-operation');
  }
  if (url !== undefined)
    writer.edit((pkg) => {
      const related = ensureHyperlinkRelationship(pkg, url!, part.name);
      if (!related) refuseWrite('The native engine refuses this URL.', 'unsupported-operation');
      next.relationshipId = related.relationshipId;
      return related.pkg;
    });
  writer.apply(part.name, {
    op: 'setHyperlinkTarget',
    linkId: target.hyperlinkId,
    ...next,
    revision,
  });
  return receipts();
}
