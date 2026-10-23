import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  withPart,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  presenceReviewFixture,
  installPresenceOwnership,
} from '../../test/support/presence-ownership';
import {
  propertyEditor,
  propertyToken,
  propertyContext,
} from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { createEditPlan } from '../diff/plan';
import { planOwnedProposalRemoval } from './owned-proposal-plan';
import { replaceOwnedProposals } from './replace-owned-proposals';
import { verifyOwnedProposalContentProofs } from './owned-proposal-content';
import { readReviewGroups } from './groups';
import { recomposeOwnedProposals } from './recompose';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { rejectAllReviewGroups } from './reject-all';
const word = (pkg: OoxmlPackage) => {
  const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw Error(reopened.reason);
  return canonicalOoxmlFingerprint(reopened.package.parts.get(pkg.mainDocumentPart)!);
};
function absent(pkg: OoxmlPackage) {
  const part = pkg.parts.get(pkg.mainDocumentPart)!,
    root = buildBlockIndex(pkg).byId.get('outer')!.controlId;
  return withPart(pkg, {
    ...part,
    root: {
      ...part.root,
      children: part.root.children.map((node) =>
        node.kind === 'body'
          ? ({
              ...node,
              children: node.children.filter((child) => child.id !== root),
            } as OoxmlElement)
          : node
      ),
    } as OoxmlElement,
  });
}
for (const remove of [false, true])
  test(`owned presence inverse ${remove} retains exact native output, detached proofs, and replacement history`, () => {
    const editor = propertyEditor(presenceReviewFixture());
    try {
      const owned = installPresenceOwnership(editor, remove);
      const pkg = editor.surface!.session.currentPackage(),
        token = propertyToken(editor);
      const before = writeOoxmlPackage(pkg);
      const planned = planOwnedProposalRemoval({ pkg, revision: token }, owned.ownership);
      if (!planned.success) throw Error(planned.error);
      expect(planned.plan.partOps.flatMap((part) => part.ops).map((op) => op.op)).toEqual(
        remove ? [] : ['removeContentControl']
      );
      expect(planned.plan.propertyDecisions).toEqual([
        { proposalId: 'presence-one', blockId: 'outer', decision: 'reject' },
      ]);
      expect(Object.isFrozen(planned.plan)).toBe(true);
      expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(before);
      const saved = readOoxmlPackage(before);
      if (!saved.ok) throw Error(saved.reason);
      expect(
        verifyOwnedProposalContentProofs(
          saved.package,
          { ...token, epoch: 'reopen' },
          owned.ownership,
          owned.contentProofs
        )
      ).toBeNull();
      const empty = createEditPlan(token, []);
      if (!empty.success) throw Error(empty.error);
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const replaced = replaceOwnedProposals(editor, owned.ownership, empty.plan, propertyContext);
      stop();
      if (!replaced.success) throw Error(replaced.error);
      expect(publications).toBe(1);
      expect(replaced.ownership).toEqual([]);
      const expected = remove ? owned.baseline : absent(owned.baseline);
      expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
      const payloads = readPayloadStore(editor.surface!.session.currentPackage());
      expect(payloads.status === 'ready' && [...payloads.entries.keys()].sort()).toEqual(
        remove ? ['inner', 'outer'] : []
      );
      const groups = readReviewGroups(editor.surface!.session.currentPackage());
      expect(groups.success && groups.groups.length).toBe(0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(before);
      expect(
        verifyOwnedProposalContentProofs(
          editor.surface!.session.currentPackage(),
          propertyToken(editor),
          owned.ownership,
          owned.contentProofs
        )
      ).toBeNull();
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
    } finally {
      editor.destroy();
    }
  }, 30000);
for (const remove of [false, true])
  test(`AI recomposition replaces owned presence ${remove} with a new native proposal and retains exact rejection`, async () => {
    const editor = propertyEditor(presenceReviewFixture());
    try {
      const owned = installPresenceOwnership(editor, remove);
      const session = editor.surface!.session;
      const source = remove ? owned.baseline : absent(owned.baseline);
      const original = await readCommittedView(session, propertyToken(editor));
      if (!original.success) throw Error(original.error);
      const markdown = original.view.projection.text.replace('Outside', 'Rewritten');
      const changed = await recomposeOwnedProposals(
        editor,
        {
          revision: propertyToken(editor),
          targetMarkdown: markdown,
          ownership: owned.ownership,
          contentProofs: owned.contentProofs,
        },
        propertyContext
      );
      if (!changed.success) throw Error(changed.error);
      expect(changed.ownership.length).toBeGreaterThan(0);
      const working = await readWorkingView(session, propertyToken(editor));
      if (!working.success) throw Error(working.error);
      expect(working.view.projection.text.trimEnd()).toBe(markdown.trimEnd());
      const active = readReviewGroups(session.currentPackage());
      if (!active.success) throw Error(active.error);
      expect(active.groups.every((group) => group.properties.length === 0)).toBe(true);
      expect(active.groups.some((group) => group.id === owned.ownership[0]!.groupId)).toBe(false);
      expect(changed.ownership.some((group) => group.groupId === owned.ownership[0]!.groupId)).toBe(
        true
      );
      const rejected = rejectAllReviewGroups(editor, propertyToken(editor), propertyContext);
      if (!rejected.success) throw Error(rejected.error);
      expect(word(session.currentPackage())).toBe(word(source));
      expect(paragraphTextOf(session.part(), session.paragraphIds().at(-1)!)).toBe('Outside');
    } finally {
      editor.destroy();
    }
  }, 30000);
