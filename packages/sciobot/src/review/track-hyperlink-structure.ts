import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { RevisionAddress } from '@docx-editor.dev/core/store';
import type { CanonicalEditPlan } from '../diff/plan';
import { fail } from './tracked-structure';
import type { TrackTextContext } from './tracked-inline';
import type { NativeHyperlinkStructureProposal, TrackHyperlinksResult } from './track-hyperlinks';
import { replaceOwnedProposals } from './replace-owned-proposals';
import { propertyAddressKey } from './property-fallback';

/** Use the same admitted native program for first proposals and later-wave replacement. */
export function trackHyperlinkStructurePlan(
  editor: DocxEditorInstance,
  plan: CanonicalEditPlan,
  context: TrackTextContext
): TrackHyperlinksResult {
  if (plan.operations.some((edit) => !['text', 'format', 'link', 'unlink'].includes(edit.kind)))
    return fail(
      'This hyperlink plan needs another tracked structural executor.',
      'unsupported-operation'
    );
  const result = replaceOwnedProposals(editor, [], plan, context);
  if (!result.success) return result;
  const proposals: NativeHyperlinkStructureProposal[] = [];
  for (const group of result.ownership) {
    const parts = new Map<string, Map<string, RevisionAddress>>();
    for (const member of group.native) {
      const addresses = parts.get(member.partName) ?? new Map<string, RevisionAddress>();
      for (const address of member.addresses)
        addresses.set(propertyAddressKey(member.partName, address), address);
      parts.set(member.partName, addresses);
    }
    for (const [partName, addresses] of parts)
      proposals.push(
        Object.freeze({
          identity: group.groupId,
          partName,
          kind: 'hyperlinkStructure',
          addresses: Object.freeze([...addresses.values()]),
        })
      );
  }
  return { success: true, revision: result.revision, proposals: Object.freeze(proposals) };
}
