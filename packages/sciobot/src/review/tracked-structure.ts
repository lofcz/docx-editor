import { withReviewGroupRevisionReservations } from './group-allocation';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  collectRevisionSites,
  withPart,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { serializeEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import { canonicalTrackedDate, type TrackTextContext } from './tracked-inline';
import { persistReviewGroups, type ReviewGroupReceipt } from './groups';

export const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const invalidXml =
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
type Surface = NonNullable<DocxEditorInstance['surface']>;

export interface TrackedStructureInput {
  readonly plan: CanonicalEditPlan;
  readonly allowGroupInitialization: boolean;
  readonly actual: RevisionToken;
  readonly surface: Surface;
  readonly part: OoxmlPart;
  readonly revision: { readonly author: string; readonly date: string };
}

/** Shared gate for tracked structural executors: context, plan copy, mount revision, and mode. */
export function readTrackedStructureInput(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext,
  subject: string
): { readonly success: true; readonly input: TrackedStructureInput } | NativeEditFailure {
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
    return fail(`Invalid tracked ${subject} context.`);
  const date = canonicalTrackedDate(data.date);
  if (!date) return fail(`Tracked ${subject} changes need a valid UTC date.`);
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
      error: `Read the current document before proposing ${subject} changes.`,
      expectedRevision: plan.revision,
      actualRevision: actual,
    };
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  return {
    success: true,
    input: {
      plan,
      allowGroupInitialization: editor.snapshot().collaborationStatus === 'inactive',
      actual,
      surface,
      part: surface.session.part(),
      revision: { author: data.author, date },
    },
  };
}

/** Every node ID of the original part, for locating nodes the transaction allocates. */
export function nodeIdsOf(part: OoxmlPart): ReadonlySet<string> {
  const ids = new Set<string>();
  const walk = (node: OoxmlPart['root'] | OoxmlPart['root']['children'][number]) => {
    ids.add(node.id);
    if (node.kind !== 'textValue') node.children.forEach(walk);
  };
  walk(part.root);
  return ids;
}

/** Run the whole transaction and both review decisions on an immutable preview. */
export function preflightTrackedStructure(
  input: TrackedStructureInput,
  ops: readonly TreeDocOp[],
  complete: (pkg: OoxmlPackage) => OoxmlPackage,
  subject: string,
  receipts: () => readonly ReviewGroupReceipt[] = () => []
): NativeEditFailure | null {
  return withReviewGroupRevisionReservations(input.surface.session.currentPackage(), () =>
    preflightTrackedStructureReserved(input, ops, complete, subject, receipts)
  );
}

function preflightTrackedStructureReserved(
  input: TrackedStructureInput,
  ops: readonly TreeDocOp[],
  complete: (pkg: OoxmlPackage) => OoxmlPackage,
  subject: string,
  receipts: () => readonly ReviewGroupReceipt[] = () => []
): NativeEditFailure | null {
  const originalSites = new Set(collectRevisionSites(input.part).map((site) => site.node.id));
  let preview = input.part;
  for (const op of ops) {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      return fail(`Native ${subject} proposal refused: ${result.reason}.`, 'unsupported-operation');
    preview = result.part;
  }
  try {
    const before = input.surface.session.currentPackage(),
      staged = complete(withPart(before, preview)),
      completed = persistReviewGroups(
        before,
        staged,
        receipts(),
        input.allowGroupInitialization
      ).parts.get(input.part.name)!;
    const sites = collectRevisionSites(completed)
      .filter((site) => !originalSites.has(site.node.id))
      .map((site) => site.node.id);
    for (const op of ['acceptAllRevisions', 'rejectAllRevisions'] as const) {
      const decision = applyTreeOp(completed, { op, siteNodeIds: sites });
      if (!decision.ok)
        return fail(
          `Native ${subject} ${op} preflight refused: ${decision.reason}.`,
          'unsupported-operation'
        );
    }
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error), 'unsupported-operation');
  }
  return null;
}

/** One gated, separately undoable native transaction with an after-allocation completion. */
export function commitTrackedStructure(
  input: TrackedStructureInput,
  ops: readonly TreeDocOp[],
  complete: (pkg: OoxmlPackage) => OoxmlPackage,
  subject: string,
  receipts: () => readonly ReviewGroupReceipt[] = () => []
): { readonly success: true; readonly revision: RevisionToken } | NativeEditFailure {
  const { surface, actual } = input,
    session = surface.session,
    beforePackage = session.currentPackage();
  try {
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [
        (pkg) =>
          persistReviewGroups(
            beforePackage,
            complete(pkg),
            receipts(),
            input.allowGroupInitialization
          ),
      ],
      true,
      true,
      'after'
    );
    if (result.rejected || !result.committed)
      return fail(
        String(result.reason ?? `Native ${subject} changes did not commit.`),
        result.reason === 'review-module-required' ? 'unsupported-operation' : 'apply-failed'
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
  };
}
