import { applyDirectListAssignments, type DirectListContext } from '../ai/direct-lists';
import {
  completeSiblingFragments,
  refuseWrite,
  type SiblingFragmentSeeds,
} from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import type { TrackedStageWriter } from './tracked-stage-writer';

type BlockSeed = Extract<SiblingFragmentSeeds['placed'][number], { kind: 'block' }>;

/** The presence group owns this complete new subtree, including its nested payloads and lists. */
export function stageSiblingApplicationBlock(
  writer: TrackedStageWriter,
  edit: Extract<CanonicalEdit, { kind: 'siblings' }>,
  seeds: SiblingFragmentSeeds,
  seed: BlockSeed,
  fresh: readonly string[],
  listContext: DirectListContext
): void {
  const within = (path: readonly number[]) => seed.path.every((index, at) => path[at] === index);
  const blockEdit = {
    ...edit,
    blocks: (edit.blocks ?? []).filter((binding) => within(binding.path)),
    atoms: (edit.atoms ?? []).filter(
      (atom) => atom.location.kind === 'new-paragraph' && within(atom.location.path)
    ),
  };
  writer.edit((pkg) => {
    const completed = completeSiblingFragments(
      pkg,
      edit.partName,
      blockEdit,
      { ...seeds, links: [], codes: [], styles: new Map(), placed: [seed] },
      fresh
    );
    const assignments = (edit.lists ?? [])
      .filter((binding) => within(binding.path))
      .map((binding) => {
        const paragraphId = completed.paragraphs.get(binding.path.join('/'));
        if (!paragraphId)
          refuseWrite('An inserted application list paragraph was not allocated.', 'apply-failed');
        return { paragraphId, before: null, target: binding.target };
      });
    return applyDirectListAssignments(completed.pkg, edit.partName, assignments, listContext);
  });
}
