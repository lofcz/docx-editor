// The evaluation collaboration fixture uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  contentControlsIn,
  findNode,
  readOoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
const attribution = { author: 'AI', date: '2026-10-24T12:00:00Z' };
const run = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const control = `<w:sdt><w:sdtPr><w:tag w:val="test:native-atom"/><w:alias w:val="Reference"/></w:sdtPr><w:sdtContent>${run(' REF ')}</w:sdtContent></w:sdt>`;
function native(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
for (const linked of [false, true])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`native atom deletion retains peer typing and complete Word histories: ${linked}/${clients}/${decision}`, async () => {
        const document = docxFixture(
          `<w:p>${run('Lead. ')}${linked ? `<w:hyperlink w:anchor="target" w:tooltip="Keep">${control}</w:hyperlink>` : control}${run(' Tail.')}</w:p>`
        );
        const options = { document, ownedAttributes: true, clientIds: clients };
        const pair = await createTwoEditorFixture(options),
          actor = pair.editors[0]!,
          peer = pair.editors[1]!;
        try {
          const original = actor.surface!.session.part(),
            paragraphId = actor.surface!.session.paragraphIds()[0]!;
          const controlId = contentControlsIn(original.root)[0]!.node.id;
          pair.pause();
          expect(
            actor.surface!.applyAutomationOps(
              () => [
                {
                  op: 'removeContentControl',
                  controlId,
                  keepContent: false,
                  revision: attribution,
                },
              ],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          const proposed = actor.surface!.session.part();
          expect(findNode(proposed, controlId)).not.toBeNull();
          expect(collectRevisionSites(proposed)).toHaveLength(1);
          expect(
            peer.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId, offset: 1, text: ' peer' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          pair.reconnect(0);
          await pair.settle();
          const check = (pending: boolean, typed: boolean) => {
            let expected = pending ? proposed : original;
            if (typed)
              expected = native(expected, {
                op: 'insertText',
                paragraphId,
                offset: 1,
                text: ' peer',
              });
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                canonicalOoxmlFingerprint(expected)
              );
          };
          check(true, true);
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(false, true);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(true, true);
          expect(peer.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(true, false);
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(false, false);
          expect(peer.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(false, true);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(true, true);
          const pendingWord = canonicalOoxmlFingerprint(actor.surface!.session.part());
          const sites = collectRevisionSites(actor.surface!.session.part()).map(
            (site) => site.node.id
          );
          expect(sites).toHaveLength(1);
          expect(
            actor.surface!.applyAutomationOps(
              () => [
                {
                  op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
                  siteNodeIds: sites,
                },
              ],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          await pair.settle();
          const decided = actor.surface!.session.part();
          expect(collectRevisionSites(decided)).toEqual([]);
          if (decision === 'reject')
            expect(canonicalOoxmlFingerprint(decided)).toBe(
              canonicalOoxmlFingerprint(
                native(original, { op: 'insertText', paragraphId, offset: 1, text: ' peer' })
              )
            );
          else expect(contentControlsIn(decided.root)).toEqual([]);
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(decided)
            );
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pendingWord);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          peer.surface!.setSelection({
            anchor: { paragraphId, offset: 1 },
            head: { paragraphId, offset: 1 },
          });
          await pair.settle();
          const geometry = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId, offset: 1 })
          );
          expect(geometry[0]).not.toBeNull();
          expect(geometry[1]).not.toBeNull();
          expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
            x: geometry[1]!.x,
            y: geometry[1]!.y,
          });
          const saved = new Uint8Array(await actor.save()),
            loaded = readOoxmlPackage(saved);
          if (!loaded.ok) throw Error(loaded.reason);
          expect(
            canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
          ).toBe(canonicalOoxmlFingerprint(decided));
          for (const shared of [false, true]) {
            const reopened = await createTwoEditorFixture({
              ...options,
              document: saved,
              clientIds: undefined,
              ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              for (const editor of reopened.editors)
                expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                  canonicalOoxmlFingerprint(decided)
                );
            } finally {
              reopened.destroy();
            }
          }
        } finally {
          pair.destroy();
        }
      });
