import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { targetDocumentInput } from '../../test/support/target-document';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { buildBlockIndex } from './block-index';
import { duplicateBlock } from './duplicate';
import { readPayloadStore } from './payload-opc-codec';
import { removeBlocks } from './remove';
import { updateAtomicBlock } from './update-atomic';

const ORIGINAL = 'molecule-middle';
const REACTION = 'CCO.CC(=O)O>>CC(=O)OCC';
async function moleculeDocument() {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Intro</w:t></w:r></w:p>');
  const plan = planTargetDocument(input, 'Intro\n\n<smiles smiles="CCO" />\n\nAfter', {
    allocateBlockId: () => ORIGINAL,
  });
  if (!plan.success) throw Error(plan.error);
  const preview = previewDirectPlan(input.pkg, plan.plan);
  if (!preview.success) throw Error(preview.error);
  return writeOoxmlPackage(preview.pkg);
}
const records = (pkg: OoxmlPackage) => {
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') throw Error('The payload store is not ready.');
  return [...store.entries.keys()].sort();
};

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const copyId of ['molecule-a', 'molecule-z'])
    test(`deleting an adopted copy preserves original metadata, peer edits, and history ${clientIds}/${copyId}`, async () => {
      const document = await moleculeDocument();
      const pair = await createTwoEditorFixture({ document, clientIds });
      try {
        const [alice, bob] = pair.editors;
        const pkg = (index: number) => pair.editors[index]!.surface!.session.currentPackage();
        const original = () => buildBlockIndex(pkg(0)).byId.get(ORIGINAL)!;
        expect(
          updateAtomicBlock(alice!, { blockId: ORIGINAL, attributes: { smiles: REACTION } }).success
        ).toBe(true);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const copied = duplicateBlock(alice!, {
          controlId: original().controlId,
          allocateBlockId: () => copyId,
        });
        if (!copied.success) throw Error(copied.error);
        await pair.settle();
        for (const index of [0, 1]) expect(records(pkg(index))).toEqual([ORIGINAL, copyId].sort());

        pair.pause();
        expect(
          removeBlocks(alice!, { controlIds: [copied.controlId], mode: 'delete' }).success
        ).toBe(true);
        expect(
          updateAtomicBlock(bob!, { blockId: ORIGINAL, attributes: { smiles: 'CCN' } }).success
        ).toBe(true);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        const check = (value: OoxmlPackage, includeCopy = false) => {
          expect(records(value)).toEqual(includeCopy ? [ORIGINAL, copyId].sort() : [ORIGINAL]);
          const blocks = buildBlockIndex(value);
          expect([...blocks.byId.keys()].sort()).toEqual(
            includeCopy ? [ORIGINAL, copyId].sort() : [ORIGINAL]
          );
          expect(blocks.byId.get(ORIGINAL)?.payload.attributes.smiles).toBe('CCN');
        };
        for (const index of [0, 1]) check(pkg(index));
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const index of [0, 1]) check(pkg(index), true);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const index of [0, 1]) {
          check(pkg(index));
          const reopened = readOoxmlPackage(new Uint8Array(await pair.editors[index]!.save()));
          if (!reopened.ok) throw Error(reopened.reason);
          check(reopened.package);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        const reopened = await createTwoEditorFixture({
          document,
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        });
        try {
          for (const editor of reopened.editors) check(editor.surface!.session.currentPackage());
        } finally {
          reopened.destroy();
        }
      } finally {
        pair.destroy();
      }
    });
