import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  acceptBlockGroup,
  groupDecisionFixture,
  proposeThreeProperties,
} from '../../test/support/group-decisions';
import { groups, groupToken } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { acceptReviewGroup } from './accept-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { paragraphOwner, propertyFallbackParagraph } from './property-fallback';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
function labelParagraph(pkg: OoxmlPackage) {
  const owner = buildBlockIndex(pkg).byId.get('middle')!;
  return propertyFallbackParagraph(pkg.parts.get(owner.partName)!, owner.controlId)!;
}
function wordText(node: OoxmlNode, mode: 'original' | 'proposed'): string {
  if (node.kind === 'textValue') return node.value;
  if (
    node.namespaceUri === WML_NAMESPACE_URI &&
    ((mode === 'original' && node.localName === 'ins') ||
      (mode === 'proposed' && node.localName === 'del'))
  )
    return '';
  return node.children.map((child) => wordText(child, mode)).join('');
}
function nativeState(pkg: OoxmlPackage, suffix: string, pending: boolean) {
  const paragraph = labelParagraph(pkg),
    part = pkg.parts.get(pkg.mainDocumentPart)!;
  expect(wordText(paragraph, 'original')).toBe((pending ? 'Old middle' : 'New middle') + suffix);
  expect(wordText(paragraph, 'proposed')).toBe('New middle' + suffix);
  const sites = collectRevisionSites(part).filter(
    (site) => paragraphOwner(part, site.node) === paragraph.id
  );
  if (pending) expect(sites.length).toBeGreaterThan(0);
  else expect(sites).toEqual([]);
}
function append(editor: DocxEditorInstance, id: string, text: string) {
  const surface = editor.surface!,
    offset = paragraphTextOf(surface.session.part(), id)!.length;
  surface.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const first of [0, 1] as const)
      test(`owned review history retains the peer decision, restores the final pending group, and preserves Word views through every undo/redo; ${clients}/${receiver}/first=${first}`, async () => {
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(groupDecisionFixture()),
          clientIds: clients,
          ownedAttributes: true,
          editor: editorOptions,
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          proposeThreeProperties(alice);
          await pair.settle();
          const before = groups(alice).map((group) => group.id),
            remaining = [before[0]!, before[2]!],
            id = labelParagraph(source(bob)).id;
          nativeState(source(alice), '', true);
          pair.pause();
          expect(acceptBlockGroup(alice).success).toBe(true);
          append(bob, id, ' peer');
          expect(acceptBlockGroup(bob).success).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          const actors = [alice, bob] as const;
          const assertAccepted = () => {
            for (const editor of pair.editors) {
              nativeState(source(editor), ' peer', false);
              expect(groups(editor).map((group) => group.id)).toEqual(remaining);
              expect(
                buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes.name
              ).toBe('New middle');
              const store = readPayloadStore(source(editor));
              expect(store.status === 'ready' && store.entries.get('middle')!.label).toBe(
                'New middle'
              );
            }
          };
          assertAccepted();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (row) => row.actorId === 'bob'
          );
          expect(remote).toBeDefined();
          expect(
            caretAt(alice.surface!.layout(), {
              paragraphId: remote!.head.nodeId,
              offset: remote!.head.offset,
            })
          ).not.toBeNull();
          for (const [actor, action] of [
            [actors[first], 'undo'],
            [actors[first], 'redo'],
            [actors[1 - first]!, 'undo'],
            [actors[1 - first]!, 'redo'],
          ] as const) {
            expect(actor.exec({ type: action }).ok).toBe(true);
            await pair.settle();
            assertAccepted();
          }
          expect(actors[first].exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          assertAccepted();
          expect(actors[1 - first]!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            nativeState(source(editor), ' peer', true);
            expect(groups(editor).map((group) => group.id)).toEqual(before);
            expect(
              buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes.name
            ).toBe('Old middle');
          }
          // No normalization step enters history: Bob's next undo removes only his earlier text.
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            nativeState(source(editor), '', true);
            expect(groups(editor).map((group) => group.id)).toEqual(before);
          }
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            nativeState(source(editor), ' peer', true);
            const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!saved.ok) throw new Error(saved.reason);
            nativeState(saved.package, ' peer', true);
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          const snapshot = Y.encodeStateAsUpdate(pair.documents[receiver]!);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(),
            roomSnapshot: snapshot,
            ownedAttributes: true,
            editor: editorOptions,
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors) {
              nativeState(source(editor), ' peer', true);
              expect(groups(editor).map((group) => group.id)).toEqual(before);
            }
          } finally {
            reopened.destroy();
          }
        } finally {
          pair.destroy();
        }
      }, 30000);

for (const mode of ['bare', 'disposed'] as const)
  test(`shared group decisions refuse before writes when owned review history is ${mode}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupDecisionFixture()),
      ownedAttributes: mode === 'disposed',
      editor: editorOptions,
    });
    try {
      const editor = pair.editors[0]!;
      proposeThreeProperties(editor);
      await pair.settle();
      if (mode === 'disposed') pair.ownership[0]!.dispose();
      const before = writeOoxmlPackage(source(editor)),
        revision = editor.surface!.session.packageRevision(),
        group = groups(editor)[1]!;
      expect(
        acceptReviewGroup(editor, group.id, groupToken(editor), propertyContext)
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(writeOoxmlPackage(source(editor))).toEqual(before);
      expect(editor.surface!.session.packageRevision()).toBe(revision);
      expect(groups(editor)).toHaveLength(3);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(groups(editor)).toEqual([]);
    } finally {
      pair.destroy();
    }
  }, 30000);

for (const receiver of [0, 1] as const)
  test(`same-group peer history stays resolved during owner deletion and restores final pending Word content; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupDecisionFixture()),
      clientIds: [41, 73],
      ownedAttributes: true,
      editor: editorOptions,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      proposeThreeProperties(alice);
      await pair.settle();
      const before = groups(alice).map((group) => group.id),
        owner = buildBlockIndex(source(alice)).byId.get('middle')!;
      pair.pause();
      expect(acceptBlockGroup(alice).success).toBe(true);
      expect(acceptBlockGroup(bob).success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      expect(
        alice.surface!.applyAutomationOps(
          () => [{ op: 'removeContentControl', controlId: owner.controlId, keepContent: false }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).committed
      ).toBe(true);
      await pair.settle();
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(buildBlockIndex(source(editor)).byId.get('middle')).toBeUndefined();
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        nativeState(source(editor), '', false);
        expect(groups(editor).map((group) => group.id)).toEqual([before[0]!, before[2]!]);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        nativeState(source(editor), '', true);
        expect(groups(editor).map((group) => group.id)).toEqual(before);
        expect(
          findNode(editor.surface!.session.part(), labelParagraph(source(editor)).id)
        ).toBeDefined();
      }
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  }, 30000);
