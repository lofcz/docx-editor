import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listParagraph } from '../../test/support/list-document';
import { readEditorRevision } from '../projection/revision-cache';
import { planParagraphText } from '../diff/paragraph-text';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { applyDirectPlan } from './apply-direct';
import { beginReviewStream } from './review-stream';

function scheduler() {
  const queue: (() => void)[] = [];
  return {
    schedule(run: () => void) {
      queue.push(run);
      return () => {
        const index = queue.indexOf(run);
        if (index >= 0) queue.splice(index, 1);
      };
    },
    frame() {
      const current = queue.splice(0);
      current.forEach((run) => run());
    },
    get size() {
      return queue.length;
    },
  };
}

test('same-frame review targets apply only the latest complete target', () => {
  const clock = scheduler(),
    calls: unknown[] = [];
  const stream = beginReviewStream({
    documentId: 'doc',
    schedule: clock.schedule,
    apply: (target: string, context) => {
      calls.push([target, context.history, context.final]);
      return target;
    },
  });
  expect(stream.push('first')).toBe(true);
  expect(stream.push('second')).toBe(true);
  expect(stream.push('final-in-frame')).toBe(true);
  expect(clock.size).toBe(1);
  clock.frame();
  expect(calls).toEqual([['final-in-frame', 'start', false]]);
  expect(stream.state).toBe('active');
});

test('later frames continue one history batch and finish flushes once', () => {
  const clock = scheduler(),
    calls: unknown[] = [];
  const stream = beginReviewStream({
    documentId: 'doc',
    schedule: clock.schedule,
    apply: (target: string, context) => {
      calls.push([target, context.history, context.final]);
      return target;
    },
  });
  stream.push('one');
  clock.frame();
  stream.push('two');
  clock.frame();
  expect(stream.finish('three')).toBe('three');
  expect(stream.finish('four')).toBeUndefined();
  expect(calls).toEqual([
    ['one', 'start', false],
    ['two', 'continue', false],
    ['three', 'continue', true],
  ]);
  expect(stream.state).toBe('ended');
});

test('cancel drops a buffered target without invoking the editor callback', () => {
  const clock = scheduler(),
    calls: string[] = [];
  const stream = beginReviewStream({
    documentId: 'doc',
    schedule: clock.schedule,
    apply: (target: string) => calls.push(target),
  });
  stream.push('discard');
  expect(stream.cancel()).toBe(true);
  clock.frame();
  expect(calls).toEqual([]);
  expect(stream.cancel()).toBe(false);
});

test('native review streaming commits the final target once and one undo restores the source', async () => {
  const pair = await createTwoEditorFixture({
    document: docxFixture(listParagraph('Original')),
    documentId: 'review-stream-native',
    ownedAttributes: true,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
  try {
    const alice = pair.editors[0]!,
      surface = alice.surface!,
      paragraphId = surface.session.paragraphIds()[0]!,
      revision = readEditorRevision(alice, 'review-stream-native', 'mount')!,
      clock = scheduler(),
      source = canonicalOoxmlFingerprint(surface.session.part());
    const targets = ['First', 'Second', 'Final'].map((text) => {
      const planned = planParagraphText(
        { revision, partName: surface.session.part().name, paragraphId, text: 'Original' },
        text
      );
      if (!planned.success) throw new Error(planned.error);
      return planned.plan;
    });
    const stream = beginReviewStream({
      documentId: 'review-stream-native',
      schedule: clock.schedule,
      apply: (target: import('../diff/plan').CanonicalEditPlan, context) =>
        applyDirectPlan(alice, target, {
          documentId: context.documentId,
          epoch: 'mount',
        }),
    });
    targets.forEach((target) => expect(stream.push(target)).toBe(true));
    clock.frame();
    expect(paragraphTextOf(surface.session.part(), paragraphId)).toBe('Final');
    expect(stream.finish()).toBeUndefined();
    await pair.settle();
    expect(paragraphTextOf(surface.session.part(), paragraphId)).toBe('Final');
    expect(paragraphTextOf(pair.editors[1]!.surface!.session.part(), paragraphId)).toBe('Final');
    const converged = await pair.savedDigests();
    expect(converged[0]).toEqual(converged[1]);
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(source);
    const saved = readOoxmlPackage(new Uint8Array(await alice.save()));
    expect(saved.ok).toBe(true);
    if (saved.ok)
      expect(
        paragraphTextOf(saved.package.parts.get(saved.package.mainDocumentPart)!, paragraphId)
      ).toBe('Original');
  } finally {
    pair.destroy();
  }
});
