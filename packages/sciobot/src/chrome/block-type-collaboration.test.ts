// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  directParagraphProperties,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { withBlockTypeStyles } from './block-styles';
import { setBlockType } from './block-type';
import { insertBlock } from './insert-block';
import { installBlockInputRules } from './block-input';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const type of ['p', 'h1', 'h2', 'h3', 'blockquote', 'hr'] as const)
    test(`${type} preserves concurrent text, peer history, geometry, deletion, reconnect, and saved convergence ${clients}`, async () => {
      const loaded = readOoxmlPackage(
        docxFixture(
          '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
        )
      );
      if (!loaded.ok) throw new Error(loaded.reason);
      const prepared = withBlockTypeStyles(loaded.package)!;
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(prepared),
        clientIds: clients,
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const id = a.session.paragraphIds()[0]!;
        const point = { paragraphId: id, offset: 5 };
        a.setSelection({ anchor: point, head: point });
        b.setSelection({
          anchor: { paragraphId: id, offset: 10 },
          head: { paragraphId: id, offset: 10 },
        });
        pair.pause();
        const result =
          type === 'hr'
            ? insertBlock(alice!, type)
            : setBlockType(alice!, type, { documentId: 'blocks', epoch: 'mount' });
        if (!result.success) throw new Error(result.error);
        b.insertPlainText(' peer');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        for (const surface of [a, b]) {
          expect(paragraphTextOf(surface.session.part(), id)).toBe('Alpha beta peer');
          expect(caretAt(surface.layout(), { paragraphId: id, offset: 12 })).not.toBeNull();
          if (type !== 'hr')
            expect(
              directParagraphProperties(surface.session.part(), id).find(
                (p) => p.localName === 'pStyle'
              )?.attributes?.val
            ).toBe(
              type === 'p' ? 'Normal' : type === 'blockquote' ? 'Quote' : `Heading${type.at(-1)}`
            );
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(id);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(a.session.part(), id)).toBe('Alpha beta peer');
        expect(
          directParagraphProperties(a.session.part(), id).find((p) => p.localName === 'pStyle')
        ).toBeUndefined();
        expect(a.session.paragraphIds()).toHaveLength(2);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
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
        for (const surface of [a, b])
          expect(paragraphTextOf(surface.session.part(), id)).toBe('Alpha beta peer!');
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        for (const editor of pair.editors) {
          const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
          expect(reopened.ok).toBe(true);
          if (!reopened.ok) throw new Error(reopened.reason);
          expect(
            paragraphTextOf(reopened.package.parts.get(reopened.package.mainDocumentPart)!, id)
          ).toBe('Alpha beta peer!');
        }
      } finally {
        pair.destroy();
      }
    });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const prefix of ['#', '--', '___'] as const)
    test(`typed block rule retains an offline insertion after its marker ${prefix}, ${clients}`, async () => {
      const loaded = readOoxmlPackage(
        docxFixture(
          `<w:p><w:r><w:t>${prefix}</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`
        )
      );
      if (!loaded.ok) throw new Error(loaded.reason);
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(withBlockTypeStyles(loaded.package)!),
        clientIds: clients,
      });
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!;
      const stop = installBlockInputRules(alice!, { documentId: 'blocks', epoch: 'mount' });
      try {
        const id = a.session.paragraphIds()[0]!,
          point = { paragraphId: id, offset: prefix.length };
        a.setSelection({ anchor: point, head: point });
        b.setSelection({ anchor: point, head: point });
        pair.pause();
        const event = new InputEvent('beforeinput', {
          bubbles: true,
          cancelable: true,
          inputType: 'insertText',
          data: prefix === '--' ? '-' : ' ',
        });
        a.hostElement().querySelector('[contenteditable="true"]')!.dispatchEvent(event);
        expect(event.defaultPrevented).toBe(true);
        b.insertPlainText('Peer');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        for (const surface of [a, b]) {
          const ids = surface.session.paragraphIds();
          expect(ids.map((value) => paragraphTextOf(surface.session.part(), value))).toEqual(
            prefix === '#' ? ['Peer', 'Tail'] : ['', 'Peer', 'Tail']
          );
          expect(
            caretAt(surface.layout(), { paragraphId: ids[prefix === '#' ? 0 : 1]!, offset: 4 })
          ).not.toBeNull();
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(
          a.session.paragraphIds().map((value) => paragraphTextOf(a.session.part(), value))
        ).toEqual([prefix + 'Peer', 'Tail']);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(
          a.session.paragraphIds().map((value) => paragraphTextOf(a.session.part(), value))
        ).toEqual(prefix === '#' ? ['Peer', 'Tail'] : ['', 'Peer', 'Tail']);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        stop();
        pair.destroy();
      }
    });
