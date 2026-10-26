import type { CanonicalEditPlan } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import { prepareApplicationPayloadStage } from './prepare-payload-stage';
import type { PropertyFallbackTarget } from './track-properties';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

/** Host-authored localized labels and fresh, explicit private proposal identities. */
export interface TrackedPayloadOptions {
  /** Fresh host identity for inserted application subtree presence ownership. */
  readonly presence?: { readonly proposalId: string };
  readonly properties?: {
    readonly intent?: import('../diff/domain-properties').DomainPropertyIntent;
    readonly proposalId: string;
    readonly fallbackTargets: readonly PropertyFallbackTarget[];
  };
  readonly atomic?: {
    readonly proposalId: string;
    readonly fallbackTargets: readonly PropertyFallbackTarget[];
  };
}

/** Stage payload records and their exact native label revisions in the caller's transaction. */
export function stageTrackedPayloads(
  writer: TrackedStageWriter,
  plan: CanonicalEditPlan,
  attribution: TrackedStageAttribution,
  options: TrackedPayloadOptions = {}
): void {
  if (
    !options ||
    typeof options !== 'object' ||
    Array.isArray(options) ||
    Object.keys(options).some((key) => !['properties', 'atomic', 'presence'].includes(key))
  )
    refuseWrite('Invalid application payload staging options.', 'invalid-input');
  for (const executor of ['properties', 'atomic'] as const) {
    const operations = plan.operations.filter(
      (edit) =>
        edit.kind === 'payload' &&
        (executor === 'atomic' ? edit.path.length === 0 : edit.path.length !== 0)
    );
    const supplied = options[executor];
    if (!operations.length) {
      if (supplied)
        refuseWrite('A fallback executor has no matching payload edits.', 'invalid-input');
      continue;
    }
    if (
      !supplied ||
      typeof supplied !== 'object' ||
      Array.isArray(supplied) ||
      Object.keys(supplied).some(
        (key) =>
          !(
            executor === 'properties'
              ? ['proposalId', 'fallbackTargets', 'intent']
              : ['proposalId', 'fallbackTargets']
          ).includes(key)
      )
    )
      refuseWrite(
        'Changed application payloads need explicit host fallback targets.',
        'unsupported-operation'
      );
    const prepared = prepareApplicationPayloadStage(
      writer.currentPackage(),
      plan.revision,
      { ...plan, operations },
      supplied.fallbackTargets,
      {
        documentId: plan.revision.documentId,
        epoch: plan.revision.epoch,
        author: attribution.author,
        date: attribution.date,
        proposalId: supplied.proposalId,
        ...(executor === 'properties' && 'intent' in supplied ? { intent: supplied.intent } : {}),
      },
      executor
    );
    if (!prepared.success) refuseWrite(prepared.error, prepared.code);
    const partName = writer.currentPackage().mainDocumentPart;
    for (const op of prepared.ops) writer.apply(partName, op);
    writer.edit((pkg) => prepared.complete(pkg).pkg);
  }
}
