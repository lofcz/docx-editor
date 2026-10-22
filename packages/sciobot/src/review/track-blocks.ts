import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  collectRevisionSites,
  contentControlsIn,
  ensureHyperlinkRelationship,
  findNode,
  paragraphTextOf,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlPackage,
  type OoxmlPart,
  type RevisionAddress,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { codeNewParagraphs } from '../diff/fragment-seeds';
import { serializeEditPlan, type CanonicalEdit, type CanonicalEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import { canonicalTrackedDate, type TrackTextContext } from './tracked-inline';
import {
  nativePropertyAddress,
  paragraphOwner,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import { newParagraphSeeds, type NewParagraphCode, type NewParagraphLink } from './new-paragraphs';
import { persistReviewGroups } from './groups';

type SiblingEdit = Extract<CanonicalEdit, { kind: 'siblings' }>;
interface Window {
  readonly edit: SiblingEdit;
  readonly paragraphs: readonly import('@docx-editor.dev/core/store').OoxmlParagraphNode[];
  readonly links: readonly NewParagraphLink[];
  readonly codes: readonly NewParagraphCode[];
}
/** Word addresses survive save/reopen. Paragraph IDs address this live editor revision only. */
export interface NativeSiblingProposal {
  readonly identity: string;
  readonly partName: string;
  readonly addresses: readonly RevisionAddress[];
  readonly insertedParagraphIds: readonly string[];
  readonly removedParagraphIds: readonly string[];
}
export type TrackBlocksResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposals: readonly NativeSiblingProposal[];
    }
  | NativeEditFailure;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const invalidXml =
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const parentNames = new Set(['body', 'tc', 'sdtContent']);
const section = (node: import('@docx-editor.dev/core/store').OoxmlNode): boolean =>
  node.kind !== 'textValue' &&
  ((node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'sectPr') ||
    node.children.some(section));
function gap(part: OoxmlPart, edit: SiblingEdit): readonly string[] | null {
  const parent = findNode(part, edit.parentId);
  if (
    !parent ||
    parent.kind === 'textValue' ||
    parent.namespaceUri !== WML_NAMESPACE_URI ||
    !parentNames.has(parent.localName)
  )
    return null;
  const before =
    edit.beforeId === null ? -1 : parent.children.findIndex((node) => node.id === edit.beforeId);
  const terminal = parent.children.findIndex(
    (node) =>
      node.kind !== 'textValue' &&
      node.namespaceUri === WML_NAMESPACE_URI &&
      node.localName === 'sectPr'
  );
  const after =
    edit.afterId === null
      ? terminal < 0
        ? parent.children.length
        : terminal
      : parent.children.findIndex((node) => node.id === edit.afterId);
  if ((edit.beforeId !== null && before < 0) || after < 0 || before >= after) return null;
  return parent.children.slice(before + 1, after).map((node) => node.id);
}
const exact = (a: readonly string[], b: readonly string[]) =>
  a.length === b.length && a.every((value, index) => value === b[index]);

/** Apply native paragraph/heading sibling gaps through one gated, separately undoable transaction. */
export function trackSiblingPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackBlocksResult {
  const copied = copyBoundedJson(context, { maxBytes: 16_384, maxNodes: 32, maxDepth: 4 });
  if (!copied.ok) return fail(copied.reason);
  const data = copied.value as unknown as TrackTextContext;
  if (
    !data ||
    typeof data.documentId !== 'string' ||
    !data.documentId ||
    typeof data.epoch !== 'string' ||
    !data.epoch ||
    typeof data.author !== 'string' ||
    !data.author.trim() ||
    data.author.length > 4096 ||
    invalidXml.test(data.author) ||
    typeof data.date !== 'string'
  )
    return fail('Invalid tracked sibling context.');
  const date = canonicalTrackedDate(data.date);
  if (!date) return fail('Tracked sibling changes need a valid UTC date.');
  let plan: CanonicalEditPlan;
  try {
    plan = JSON.parse(serializeEditPlan(input)) as CanonicalEditPlan;
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error));
  }
  const actual = readEditorRevision(editor, data.documentId, data.epoch),
    surface = editor.surface;
  if (!actual || !surface) return fail('The native editor is not attached.', 'document-not-ready');
  if (!sameRevision(actual, plan.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before proposing sibling changes.',
      expectedRevision: plan.revision,
      actualRevision: actual,
    };
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  const session = surface.session,
    part = session.part(),
    windows: Window[] = [],
    removed = new Set<string>(),
    locations = new Set<string>();
  const initialSites = collectRevisionSites(part);
  const originalSites = new Map(
    initialSites.map((site) => [site.node.id, JSON.stringify(site.node)])
  );
  const pendingParagraphs = new Set(initialSites.map((site) => paragraphOwner(part, site.node)));
  for (const edit of plan.operations) {
    if (
      edit.kind !== 'siblings' ||
      edit.partName !== part.name ||
      edit.lists?.some((binding) => binding.target.kind !== 'none')
    )
      return fail(
        'This executor accepts body paragraph/heading sibling plans; list and resource fragments keep their own executors.',
        'unsupported-operation'
      );
    const ids = gap(part, edit);
    if (!ids || !exact(ids, edit.removedIds))
      return fail(
        'The original native sibling gap no longer matches the plan.',
        'ambiguous-target'
      );
    const location = JSON.stringify([edit.parentId, edit.beforeId, edit.afterId]);
    if (locations.has(location)) return fail('Planned sibling gaps overlap.', 'ambiguous-target');
    locations.add(location);
    const protection = propertyControlRefusal(part, edit.parentId);
    if (protection) return fail(protection, 'unsupported-operation');
    for (const id of edit.removedIds) {
      if (removed.has(id))
        return fail('A paragraph cannot belong to two structural gaps.', 'ambiguous-target');
      removed.add(id);
      const paragraph = findNode(part, id);
      if (paragraph?.kind !== 'paragraph' || section(paragraph))
        return fail(
          'Only complete native paragraphs without section marks can be removed here.',
          'unsupported-operation'
        );
      const controls = [
        propertyControlRefusal(part, id),
        ...contentControlsIn(paragraph).map((entry) => propertyControlRefusal(part, entry.node.id)),
      ];
      if (controls.some(Boolean))
        return fail(
          'Protected paragraph content cannot enter this proposal.',
          'unsupported-operation'
        );
      if (pendingParagraphs.has(id))
        return fail(
          'Deleting an earlier pending paragraph proposal needs explicit ownership reconciliation.',
          'unsupported-operation'
        );
    }
    const seeds = newParagraphSeeds(edit.inserted, { code: true });
    if (!seeds.success) return seeds;
    windows.push({ edit, paragraphs: seeds.paragraphs, links: seeds.links, codes: seeds.codes });
  }
  if (
    windows.some(
      ({ edit }) =>
        (edit.beforeId !== null && removed.has(edit.beforeId)) ||
        (edit.afterId !== null && removed.has(edit.afterId))
    )
  )
    return fail('A sibling anchor cannot be removed by another planned gap.', 'ambiguous-target');
  const ops: TreeDocOp[] = [],
    revision = { author: data.author, date };
  for (const { edit, paragraphs } of windows) {
    for (const id of edit.removedIds) {
      const text = paragraphTextOf(part, id)!;
      if (text.length)
        ops.push({ op: 'deleteText', paragraphId: id, start: 0, end: text.length, revision });
      ops.push({ op: 'setParagraphMarkRevision', paragraphId: id, kind: 'del', revision });
    }
    if (paragraphs.length)
      ops.push({
        op: 'insertParagraphs',
        parentId: edit.parentId,
        beforeId: edit.afterId,
        paragraphs,
        revision,
      });
  }
  if (!ops.length)
    return { success: true, revision: Object.freeze(actual), proposals: Object.freeze([]) };
  let preview = part;
  for (const op of ops) {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      return fail(`Native sibling proposal refused: ${result.reason}.`, 'unsupported-operation');
    preview = result.part;
  }
  const previewSites = collectRevisionSites(preview)
    .filter((site) => !originalSites.has(site.node.id))
    .map((site) => site.node.id);
  for (const op of ['acceptAllRevisions', 'rejectAllRevisions'] as const) {
    const decision = applyTreeOp(preview, { op, siteNodeIds: previewSites });
    if (!decision.ok)
      return fail(
        `Native sibling ${op} preflight refused: ${decision.reason}.`,
        'unsupported-operation'
      );
  }
  let proposals: readonly NativeSiblingProposal[] = [];
  const beforePackage = surface.session.currentPackage();
  const complete = (pkg: OoxmlPackage): OoxmlPackage => {
    let working = pkg;
    let current = pkg.parts.get(part.name)!;
    const ownership = new Map<string, number>(),
      inserted: string[][] = [];
    const originalIds = new Set<string>();
    const walk = (node: import('@docx-editor.dev/core/store').OoxmlNode) => {
      originalIds.add(node.id);
      if (node.kind !== 'textValue') node.children.forEach(walk);
    };
    walk(part.root);
    windows.forEach(({ edit, paragraphs }, index) => {
      const parent = findNode(current, edit.parentId) as OoxmlElement;
      const before =
        edit.beforeId === null
          ? -1
          : parent.children.findIndex((node) => node.id === edit.beforeId);
      const after =
        edit.afterId === null
          ? parent.children.length
          : parent.children.findIndex((node) => node.id === edit.afterId);
      const fresh = parent.children
        .slice(before + 1, after)
        .filter((node) => node.kind === 'paragraph' && !originalIds.has(node.id))
        .map((node) => node.id);
      if (fresh.length !== paragraphs.length)
        throw new Error('Native sibling allocation did not match the planned gap.');
      inserted.push(fresh);
      for (const id of [...edit.removedIds, ...fresh]) {
        if (ownership.has(id)) throw new Error('Native sibling ownership overlaps.');
        ownership.set(id, index);
      }
    });
    windows.forEach(({ links }, index) => {
      for (const link of links) {
        const paragraphId = inserted[index]![link.paragraphIndex]!;
        const target: { relationshipId: string } | { anchor: string } = (() => {
          if (link.target.kind === 'bookmark') return { anchor: link.target.anchor };
          const resource = ensureHyperlinkRelationship(working, link.target.url, part.name);
          if (!resource) throw new Error('A new native hyperlink target was refused.');
          working = resource.pkg;
          return { relationshipId: resource.relationshipId };
        })();
        const wrapped = applyTreeOp(current, {
          op: 'insertHyperlink',
          paragraphId,
          start: link.start,
          end: link.end,
          ...target,
          ...(link.title !== null ? { tooltip: link.title } : {}),
        });
        if (!wrapped.ok) throw new Error(`A new native hyperlink was refused: ${wrapped.reason}.`);
        current = wrapped.part;
      }
    });
    working = withPart(working, current);
    windows.forEach(({ codes }, index) => {
      working = codeNewParagraphs(working, part.name, inserted[index]!, codes);
    });
    current = working.parts.get(part.name)!;
    const sites = collectRevisionSites(current).filter((site) => !originalSites.has(site.node.id));
    const attributed = sites.map((site) => {
      const owner = paragraphOwner(current, site.node),
        index = owner === null ? undefined : ownership.get(owner);
      if (index === undefined) throw new Error('A new native revision escaped its structural gap.');
      if (!['ins', 'del'].includes(site.node.localName))
        throw new Error('Unexpected native sibling revision type.');
      return { nodeId: site.node.id, renumber: `sibling:${index}:${site.node.localName}` };
    });
    const renamed = applyTreeOp(current, {
      op: 'setRevisionAttribution',
      revision,
      sites: attributed,
    });
    if (!renamed.ok) throw new Error(`Native sibling attribution refused: ${renamed.reason}.`);
    current = renamed.part;
    for (const [id, before] of originalSites)
      if (JSON.stringify(findNode(current, id)) !== before)
        throw new Error('A sibling proposal changed an unrelated pending revision.');
    const addresses = windows.map(() => new Map<string, RevisionAddress>());
    for (const site of collectRevisionSites(current)) {
      if (originalSites.has(site.node.id)) continue;
      const owner = paragraphOwner(current, site.node),
        index = owner === null ? undefined : ownership.get(owner),
        address = nativePropertyAddress(site.node);
      if (index === undefined || !address)
        throw new Error('Missing native sibling review receipt.');
      addresses[index]!.set(propertyAddressKey(part.name, address), Object.freeze({ ...address }));
    }
    proposals = Object.freeze(
      windows.map(({ edit, paragraphs }, index) => {
        const group = [...addresses[index]!.values()];
        if (
          group.length !==
          Number(Boolean(paragraphs.length)) + Number(Boolean(edit.removedIds.length))
        )
          throw new Error(
            'A structural gap did not retain distinct native insertion/deletion addresses.'
          );
        return Object.freeze({
          identity: JSON.stringify([
            part.name,
            group.map((address) => propertyAddressKey(part.name, address)).sort(),
          ]),
          partName: part.name,
          addresses: Object.freeze(group),
          insertedParagraphIds: Object.freeze(inserted[index]!),
          removedParagraphIds: Object.freeze([...edit.removedIds]),
        });
      })
    );
    return persistReviewGroups(
      beforePackage,
      withPart(working, current),
      proposals,
      editor.snapshot().collaborationStatus === 'inactive'
    );
  };
  try {
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [complete],
      true,
      true,
      'after'
    );
    if (result.rejected || !result.committed)
      return fail(
        String(result.reason ?? 'Native sibling changes did not commit.'),
        'apply-failed'
      );
  } catch (error) {
    const changed = session.packageRevision() !== actual.revision;
    return {
      success: false,
      code: 'apply-failed',
      error: `${changed ? 'The document changed during a failed native call. Re-read before retrying. ' : ''}${error instanceof Error ? error.message : String(error)}`,
      expectedRevision: actual,
      actualRevision: { ...actual, revision: session.packageRevision() },
    };
  }
  return {
    success: true,
    revision: Object.freeze({ ...actual, revision: session.packageRevision() }),
    proposals,
  };
}
