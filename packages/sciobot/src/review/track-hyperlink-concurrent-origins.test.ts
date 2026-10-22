import { rejectReviewGroup } from './reject-group';
import { acceptReviewGroup } from './accept-group';
import { groupToken } from '../../test/support/review-groups';
// The existing test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  validateOoxmlPart,
  serializeOoxmlPart,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { runWithTransactionActor } from '@docx-editor.dev/core/store/package/actor-scoped-ids';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  document_,
  context,
  plan,
  fingerprint,
  projection,
  ast,
} from '../../test/support/link-structure-execution';
import { propertyEditor } from '../../test/support/property-execution';
import { prepareReviewGroupDocument } from '../native/prepare-review-groups';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';
import { applyDirectPlan } from '../ai/apply-direct';
import { RUN_ORIGIN_NS } from '@docx-editor.dev/core/store/store/run-origin-metadata';
const one = '[one](https://example.org/one)',
  two = '[two](https://example.org/two)';
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`concurrent proposal provenance ${clientIds}/${receiver}/${action}`, async () => {
        const prepared = prepareReviewGroupDocument(document_());
        if (!prepared.success) throw Error(prepared.error);
        const pair = await createTwoEditorFixture({
          document: prepared.document,
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          clientIds,
        });
        const read = readOoxmlPackage(prepared.document);
        if (!read.ok) throw Error(read.reason);
        const golden = propertyEditor(read.package);
        try {
          const [alice, bob] = pair.editors;
          pair.pause();
          const a = trackHyperlinkPlan(
            alice!,
            await plan(alice!, `Alpha end ${one}\n\nTail`),
            context
          );
          if (!a.success) throw Error(a.error);
          const b = trackHyperlinkPlan(bob!, await plan(bob!, `Alpha end ${two}\n\nTail`), context);
          if (!b.success) throw Error(b.error);
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(validateOoxmlPart(editor.surface!.session.part())).toEqual({ ok: true });
            expect(ast((await projection(editor, true)).text)).toBe(ast('Alpha end\n\nTail'));
          }
          expect(fingerprint(alice!)).toBe(fingerprint(bob!));
          const proposed = (await projection(alice!)).text;
          const choices = [`Alpha end ${one} ${two}\n\nTail`, `Alpha end ${two} ${one}\n\nTail`];
          expect(choices.map(ast)).toContain(ast(proposed));
          if (action === 'accept') {
            const aFirst = proposed.indexOf('[one]') < proposed.indexOf('[two]');
            const first = `Alpha end ${aFirst ? one : two}\n\nTail`,
              actor = aFirst ? 'alice' : 'bob';
            const p1 = await plan(golden, first);
            const d1 = runWithTransactionActor(actor, () => applyDirectPlan(golden, p1, context));
            if (!d1.success) throw Error(d1.error);
            const p2 = await plan(golden, proposed);
            const d2 = runWithTransactionActor(aFirst ? 'bob' : 'alice', () =>
              applyDirectPlan(golden, p2, context)
            );
            if (!d2.success) throw Error(d2.error);
          }
          expect(alice!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(validateOoxmlPart(editor.surface!.session.part())).toEqual({ ok: true });
            expect(fingerprint(editor)).toBe(fingerprint(golden));
            expect(serializeOoxmlPart(editor.surface!.session.part())).not.toContain(RUN_ORIGIN_NS);
          }
        } finally {
          pair.destroy();
          golden.destroy();
        }
      }, 30000);

for (const order of [
  [0, 1],
  [1, 0],
] as const)
  for (const actions of [
    ['accept', 'accept'],
    ['accept', 'reject'],
    ['reject', 'accept'],
    ['reject', 'reject'],
  ] as const)
    test(`concurrent proposals decide independently in order ${order}/${actions}`, async () => {
      const prepared = prepareReviewGroupDocument(document_());
      if (!prepared.success) throw Error(prepared.error);
      const pair = await createTwoEditorFixture({
        document: prepared.document,
        documentId: context.documentId,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        clientIds: [41, 73],
      });
      const read = readOoxmlPackage(prepared.document);
      if (!read.ok) throw Error(read.reason);
      const golden = propertyEditor(read.package);
      try {
        const [alice, bob] = pair.editors;
        pair.pause();
        const a = trackHyperlinkPlan(
          alice!,
          await plan(alice!, `Alpha end ${one}\n\nTail`),
          context
        );
        const b = trackHyperlinkPlan(bob!, await plan(bob!, `Alpha end ${two}\n\nTail`), context);
        if (!a.success || !b.success) throw Error('proposal failed');
        pair.reconnect(0);
        await pair.settle();
        const proposed = (await projection(alice!)).text;
        const aFirst = proposed.indexOf('[one]') < proposed.indexOf('[two]');
        const kept = (aFirst ? [0, 1] : [1, 0]).filter((index) => actions[index] === 'accept');
        let target = 'Alpha end';
        for (const index of kept) {
          target += ` ${index === 0 ? one : two}`;
          const direct = await plan(golden, `${target}\n\nTail`);
          expect(
            runWithTransactionActor(index === 0 ? 'alice' : 'bob', () =>
              applyDirectPlan(golden, direct, context)
            ).success
          ).toBe(true);
        }
        const proposals = [a, b];
        for (const index of order) {
          const before = fingerprint(alice!);
          expect(
            (actions[index] === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              alice!,
              proposals[index]!.proposals[0]!.identity,
              groupToken(alice!),
              context
            ).success
          ).toBe(true);
          await pair.settle();
          const after = fingerprint(alice!);
          expect(fingerprint(bob!)).toBe(after);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(before);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(after);
          if (index === order[0]) {
            const read = readOoxmlPackage(new Uint8Array(await alice!.save()));
            if (!read.ok) throw Error(read.reason);
            const reopened = propertyEditor(read.package);
            try {
              const remaining = order[1];
              const decide =
                actions[remaining] === 'accept' ? acceptReviewGroup : rejectReviewGroup;
              expect(
                decide(
                  reopened,
                  proposals[remaining]!.proposals[0]!.identity,
                  groupToken(reopened),
                  context
                ).success
              ).toBe(true);
              expect(fingerprint(reopened)).toBe(fingerprint(golden));
            } finally {
              reopened.destroy();
            }
          }
        }
        for (const editor of pair.editors) expect(fingerprint(editor)).toBe(fingerprint(golden));
      } finally {
        pair.destroy();
        golden.destroy();
      }
    }, 30000);
