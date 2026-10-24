import { planTargetDocument, type TargetDocumentOptions } from '../diff/plan-target';
import type { TargetSourceInput } from '../diff/target-source';

/** Check native equivalence without invoking host allocation or payload services. */
export function isCurrentNativeTarget(
  input: TargetSourceInput,
  markdown: string,
  definitions?: TargetDocumentOptions['definitions']
): boolean {
  const needsHost = (): never => {
    throw Error('A changed target needs its normal planning services.');
  };
  try {
    const result = planTargetDocument(input, markdown, {
      definitions,
      allocateBlockId: needsHost,
      newBlockPayload: needsHost,
      headingStyle: needsHost,
      quoteStyle: needsHost,
    });
    return result.success && !result.plan.operations.length;
  } catch {
    return false;
  }
}
