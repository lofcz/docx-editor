import { prepareNativeDocument } from '../../src/native/prepare-document';
import { MDX_BLOCK } from '../../src/projection/mdx';
import type { TargetDocumentOptions } from '../../src/diff/plan-target';
import type { JsonObject } from '../../src/data/json';
export const applicationFragmentTarget = (name: string) =>
  `Outside\n\n<phase type="evocation">\n<activity name="${name}" duration="5">\nA **rich** [link](https://example.org).\n\n- One\n\n- Two\n</activity>\n</phase>\n\n<lesson_visual kind="image" src="vfs://diagram.png" alt="Diagram" />\n\n<BudgetCard title="${name}" config={{ rate: 2 }}>\n\n<Part amount={3} />\n\n</BudgetCard>`;
export const applicationFragmentOptions = (prefix: string): TargetDocumentOptions => {
  let count = 0;
  return {
    allocateBlockId: () => `${prefix}-${++count}`,
    newBlockPayload: (ast, id) =>
      ast.name === 'BudgetCard'
        ? {
            label: 'Budget preview',
            payload: {
              version: 1,
              id,
              type: MDX_BLOCK,
              attributes: { host: 'react' },
              mdast: ast as JsonObject,
              extension: { keep: [null, 2] },
            },
          }
        : null,
  };
};

export function preparedApplicationDocument(document: Uint8Array): Uint8Array {
  const prepared = prepareNativeDocument(document);
  if (!prepared.success) throw Error(prepared.error);
  return prepared.document;
}
