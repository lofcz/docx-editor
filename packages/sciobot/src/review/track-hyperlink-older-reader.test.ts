// The existing test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import {
  readOoxmlPackage,
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  collectRevisionSites,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import {
  originNodes,
  hasOriginMetadata,
} from '@docx-editor.dev/core/store/store/run-origin-metadata';
import { runWithTransactionActor } from '@docx-editor.dev/core/store/package/actor-scoped-ids';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { document_, context, plan, fingerprint } from '../../test/support/link-structure-execution';
import { propertyEditor } from '../../test/support/property-execution';
import { prepareReviewGroupDocument } from '../native/prepare-review-groups';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';
import { applyDirectPlan } from '../ai/apply-direct';
for (const fixture of ['legacy', 'dynamic', 'layered'])
  for (const priorAction of ['accept', 'reject'] as const) {
    const prefix = fixture === 'legacy' ? '' : fixture + '-';
    const data = await Bun.file(
      new URL(
        `../../../core/src/store/__tests__/fixtures/run-origins/${prefix}installed-reader-output.json`,
        import.meta.url
      )
    ).json();
    const bytes = unzipSync(document_());
    bytes['word/document.xml'] = strToU8(
      priorAction === 'accept' ? data.acceptedXml : data.rejectedXml
    );
    const prepared = prepareReviewGroupDocument(zipSync(bytes));
    if (!prepared.success) throw Error(prepared.error);
    const document = prepared.document;
    for (const clientIds of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const receiver of [0, 1] as const)
        for (const inside of [false, true])
          for (const action of (inside ? ['accept'] : ['accept', 'reject']) as readonly (
            | 'accept'
            | 'reject'
          )[])
            test(`later link proposal after ${fixture} older-core ${priorAction} ${clientIds}/${receiver}/${action}/${inside ? 'inside-copy' : 'suffix'}`, async () => {
              const pair = await createTwoEditorFixture({
                document,
                documentId: context.documentId,
                ownedAttributes: true,
                editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
                clientIds,
              });
              const read = readOoxmlPackage(document);
              if (!read.ok) throw Error(read.reason);
              const golden = propertyEditor(read.package);
              const append = (editor: ReturnType<typeof propertyEditor>) => {
                const id = editor.surface!.session.paragraphIds()[0]!,
                  p = findNode(editor.surface!.session.part(), id)!;
                if (p.kind !== 'paragraph') throw Error('paragraph');
                const offsets = paragraphOffsetIndex(p);
                const edits: TreeDocOp[] = [
                  {
                    op: 'insertText',
                    paragraphId: id,
                    offset: paragraphOffsetIndex(p).length,
                    text: '!',
                  },
                ];
                if (inside) {
                  const link = originNodes(p).find((node) => node.kind === 'hyperlink')!;
                  const span = offsets.spanOf(link)!;
                  expect(span.end - span.start).toBe(4);
                  edits.push({
                    op: 'insertText',
                    paragraphId: id,
                    offset: span.start + 2,
                    text: 'P1',
                  });
                }
                return editor.surface!.applyAutomationOps(
                  () => edits,
                  { kind: 'body' },
                  [],
                  false,
                  true
                );
              };
              try {
                const [alice, bob] = pair.editors;
                const suffix =
                    priorAction === 'accept'
                      ? fixture === 'layered'
                        ? 'YXB'
                        : 'XB'
                      : fixture === 'layered'
                        ? 'XB'
                        : 'B',
                  target = `A [next](https://example.org/later) ${suffix}`;
                const source = alice!.surface!.session.part();
                expect(collectRevisionSites(source)).toHaveLength(0);
                expect(originNodes(source.root).some(hasOriginMetadata)).toBe(true);
                expect(paragraphTextOf(source, alice!.surface!.session.paragraphIds()[0]!)).toBe(
                  'A' + suffix
                );
                const result = trackHyperlinkPlan(alice!, await plan(alice!, target), context);
                if (!result.success) throw Error(result.error);
                await pair.settle();
                pair.pause();
                expect(append(bob!)).toMatchObject({ committed: true, rejected: false });
                pair.reconnect(receiver);
                await pair.settle();
                const pending = fingerprint(alice!);
                if (action === 'accept') {
                  const p = await plan(golden, target);
                  const direct = runWithTransactionActor('alice', () =>
                    applyDirectPlan(golden, p, context)
                  );
                  if (!direct.success) throw Error(direct.error);
                }
                const withoutPeer = fingerprint(golden);
                expect(append(golden)).toMatchObject({ committed: true, rejected: false });
                const withPeer = fingerprint(golden);
                expect(alice!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
                await pair.settle();
                for (const editor of pair.editors) expect(fingerprint(editor)).toBe(withPeer);
                expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
                await pair.settle();
                for (const editor of pair.editors) expect(fingerprint(editor)).toBe(withoutPeer);
                expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
                await pair.settle();
                for (const editor of pair.editors) expect(fingerprint(editor)).toBe(withPeer);
                expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
                await pair.settle();
                for (const editor of pair.editors) expect(fingerprint(editor)).toBe(pending);
                expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
                await pair.settle();
                for (const editor of pair.editors) expect(fingerprint(editor)).toBe(withPeer);
              } finally {
                pair.destroy();
                golden.destroy();
              }
            }, 30000);
  }
