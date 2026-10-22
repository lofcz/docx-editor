// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { setSurfaceInlineWidgetMetrics } from '@docx-editor.dev/core/editor';
import { updateSurfaceMeasurement } from '../../../core/src/editor/surface-measurement';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { inlineWidgetDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { writeBlockPayload } from '../blocks/payload-part';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`native widget deletion preserves peer payloads and text through history and reconnect ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: inlineWidgetDocument(),
        clientIds: clients,
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        for (const [peer, surface] of [a, b].entries()) {
          // Fixed advances make the remote caret oracle independent of document font defaults.
          updateSurfaceMeasurement(surface, {
            measurer: {
              measure: (text) => text.length * 6,
              lineMetrics: () => ({ height: 14, baseline: 11 }),
            },
            producer: 'inline-atom-test',
          });
          setSurfaceInlineWidgetMetrics(surface, {
            measure: () => ({ width: peer === 0 ? 64 : 40, ascent: 20, descent: 4 }),
            revision: '1',
          });
        }
        const id = a.session.paragraphIds()[0]!;
        pair.pause();
        const point = { paragraphId: id, offset: 13 };
        a.setSelection({ anchor: point, head: point });
        a.deleteBackward();
        const block = buildBlockIndex(b.session.currentPackage()).byId.get('source-1')!;
        expect(
          writeBlockPayload(bob!, {
            controlId: block.controlId,
            payload: { ...block.payload, attributes: { name: 'Peer' } },
            label: 'Source',
            fallbackOps: [],
          }).success
        ).toBe(true);
        const end = { paragraphId: id, offset: 19 };
        b.setSelection({ anchor: end, head: end });
        b.insertPlainText(' Bob');
        b.flushPendingInput();
        pair.reconnect(receiver);
        await pair.settle();
        for (const surface of [a, b]) {
          expect(paragraphTextOf(surface.session.part(), id)).toBe('Before  after Bob');
          expect(buildBlockIndex(surface.session.currentPackage()).byId.has('source-1')).toBe(
            false
          );
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const surface of [a, b]) {
          expect(paragraphTextOf(surface.session.part(), id)).toBe('Before Source after Bob');
          expect(
            buildBlockIndex(surface.session.currentPackage()).byId.get('source-1')!.payload
              .attributes
          ).toEqual({ name: 'Peer' });
        }
        const last = { paragraphId: id, offset: 23 };
        b.setSelection({ anchor: last, head: last });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote.head.offset).toBe(23);
        expect(
          caretAt(
            a.layout(),
            { paragraphId: remote.head.nodeId, offset: remote.head.offset },
            a.textMeasurer?.()
          )!.x
        ).toBe(150);
        expect(caretAt(b.layout(), last, b.textMeasurer?.())!.x).toBe(132);
        const restored = await pair.savedDigests();
        expect(restored[0]).toEqual(restored[1]);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const surface of [a, b])
          expect(paragraphTextOf(surface.session.part(), id)).toBe('Before  after Bob');
        const deleted = await pair.savedDigests();
        expect(deleted[0]).toEqual(deleted[1]);
      } finally {
        pair.destroy();
      }
    });
