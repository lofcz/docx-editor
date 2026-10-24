// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import {
  directParagraphProperties,
  paragraphTextOf,
  readOoxmlPackage,
  canonicalOoxmlFingerprint,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  PARAGRAPH_ALIGNMENTS,
  setParagraphAlignment,
  changeParagraphIndent,
} from './paragraph-position';
const choices = [...PARAGRAPH_ALIGNMENTS, 'increase', 'decrease'] as const;

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const value of choices)
    test(`paragraph position ${value} preserves offline peer edits, history, exact neighbor and saved convergence ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: docxFixture(
          '<w:p><w:pPr><w:spacing w:line="251" w:lineRule="exact"/></w:pPr><w:r><w:t>Imported</w:t></w:r></w:p><w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p/>'
        ),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const ids = a.session.paragraphIds(),
          id = ids[1]!;
        const imported = directParagraphProperties(a.session.part(), ids[0]!);
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 5 },
        };
        a.setSelection(selection);
        const originalProperties = directParagraphProperties(a.session.part(), id);
        const apply = () =>
          value === 'increase' || value === 'decrease'
            ? changeParagraphIndent(alice!, value, { documentId: 'position', epoch: 'mount' })
            : setParagraphAlignment(alice!, value, { documentId: 'position', epoch: 'mount' });
        const expected: Record<string, string> =
          value === 'increase'
            ? { left: '360' }
            : value === 'decrease'
              ? { left: '0' }
              : { val: value === 'justify' ? 'both' : value };
        const property = value === 'increase' || value === 'decrease' ? 'ind' : 'jc';
        b.setSelection({
          anchor: { paragraphId: id, offset: 10 },
          head: { paragraphId: id, offset: 10 },
        });
        pair.pause();
        expect(apply().success).toBe(true);
        b.insertPlainText(' peer');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        for (const editor of pair.editors) {
          const s = editor.surface!;
          s.setSelection(selection);
          expect(
            directParagraphProperties(s.session.part(), id).find((p) => p.localName === property)
              ?.attributes
          ).toEqual(expected);
          expect(paragraphTextOf(s.session.part(), id)).toBe('Alpha beta peer');
          expect(directParagraphProperties(s.session.part(), ids[0]!)).toEqual(imported);
          expect(caretAt(s.layout(), { paragraphId: id, offset: 12 })).not.toBeNull();
          expect(
            selectionRects(s.layout(), selection, s.session.paragraphIds()).length
          ).toBeGreaterThan(0);
        }
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(directParagraphProperties(a.session.part(), id)).toEqual(originalProperties);

        expect(paragraphTextOf(a.session.part(), id)).toBe('Alpha beta peer');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(a.session.part(), id)).toBe('Alpha beta');
        expect(
          directParagraphProperties(a.session.part(), id).find((p) => p.localName === property)
            ?.attributes
        ).toEqual(expected);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        pair.pause();
        b.setSelection({
          anchor: { paragraphId: id, offset: 15 },
          head: { paragraphId: id, offset: 15 },
        });
        b.insertPlainText('!');
        b.flushPendingInput();
        expect(
          a.applyAutomationOps(
            () => [{ op: 'deleteBlock', blockId: id }],
            { kind: 'body' },
            undefined,
            false,
            true
          ).rejected
        ).toBe(false);
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Alpha beta peer!');
          expect(directParagraphProperties(editor.surface!.session.part(), ids[0]!)).toEqual(
            imported
          );
          const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!reopened.ok) throw new Error(reopened.reason);
          expect(
            canonicalOoxmlFingerprint(
              reopened.package.parts.get(reopened.package.mainDocumentPart)!
            )
          ).toEqual(canonicalOoxmlFingerprint(editor.surface!.session.part()));
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });
