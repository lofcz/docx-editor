import { expect, test } from 'bun:test';
import { createBrowserAutomationHost } from '@docx-editor.dev/core/editor';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../../../e2e/support/sciobot-collaboration';
import { fixture } from '../../../../examples/react-blocks/src/document';

test('two real sessions synchronize, reconnect, undo, and delete a paragraph during peer editing', async () => {
  const pair = await createTwoEditorFixture({
    document: fixture(),
    documentId: 'adapter-harness-smoke',
  });
  const paragraph = (index: number, text: string) =>
    [...pair.containers[index]!.querySelectorAll<HTMLElement>('[data-paragraph-id]')].find((node) =>
      node.textContent?.includes(text)
    );
  const read = (index: number, id: string) =>
    paragraphTextOf(pair.editors[index]!.surface!.session.part(), id);
  const type = (index: number, id: string, text: string) => {
    pair.editors[index]!.surface!.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    pair.editors[index]!.surface!.insertPlainText(text);
    pair.flush();
  };
  try {
    const id = paragraph(0, 'Plan an experiment.')!.dataset.paragraphId!;
    type(0, id, 'Alice: ');
    expect(read(1, id)).toBe(read(0, id));
    expect(pair.rooms[1]!.session.remoteSelections().length).toBeGreaterThan(0);
    pair.pause();
    type(0, id, 'A ');
    type(1, id, 'B ');
    pair.reconnect();
    expect(read(0, id)).toBe(read(1, id));
    expect(read(0, id)).toContain('A ');
    expect(read(0, id)).toContain('B ');
    const combined = read(0, id);
    expect(pair.editors[0]!.exec({ type: 'undo' }).ok).toBe(true);
    pair.flush();
    expect(read(0, id)).toBe(read(1, id));
    expect(read(0, id)).not.toBe(combined);
    expect(read(0, id)).toContain('B ');
    expect(pair.editors[0]!.exec({ type: 'redo' }).ok).toBe(true);
    pair.flush();
    expect(read(0, id)).toBe(combined);
    expect(read(1, id)).toBe(combined);

    const target = paragraph(1, 'Activity step 3.')!.dataset.paragraphId!;
    pair.pause();
    type(1, target, 'Bob is editing here. ');
    const host = createBrowserAutomationHost(pair.editors[0]!);
    try {
      const document = host.execute({ operations: [{ op: 'getDocument' }] }).results[0]!;
      if (document.status !== 'ok' || document.value.kind !== 'handle')
        throw new Error('No document');
      const body = host.execute({
        operations: [{ op: 'getBody', document: document.value.handle }],
      }).results[0]!;
      if (body.status !== 'ok' || body.value.kind !== 'handle') throw new Error('No body');
      const found = host.execute({
        operations: [
          { op: 'search', scope: { body: body.value.handle }, text: 'Activity step 3.' },
        ],
      }).results[0]!;
      if (found.status !== 'ok' || found.value.kind !== 'spans') throw new Error('No target');
      expect(
        host.execute({
          operations: [{ op: 'deleteParagraph', paragraph: found.value.spans[0]!.start.paragraph }],
        }).ok
      ).toBe(true);
    } finally {
      host.dispose();
    }
    pair.flush();
    pair.reconnect();
    await pair.settle();
    expect(read(0, target)).toBeNull();
    expect(read(1, target)).toBeNull();
    expect(Boolean(paragraph(0, 'Activity step 3.'))).toBe(false);
    expect(Boolean(paragraph(1, 'Activity step 3.'))).toBe(false);
    const [alice, bob] = await pair.savedDigests();
    expect(alice).toEqual(bob);
  } finally {
    pair.destroy();
    pair.destroy();
  }
}, 30000);
