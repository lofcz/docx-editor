import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { zipDocument } from '../../../pro/src/collaboration/__tests__/document-peer-support';

const bytes = zipDocument(
  '<w:p><w:r><w:t>Document text</w:t></w:r></w:p><w:p><w:r><w:t>Other text</w:t></w:r></w:p>'
);
function draft(document: Document, owner: string) {
  const wrapper = document.createElement('div');
  wrapper.setAttribute('data-docx-host-ui', '');
  wrapper.setAttribute('data-docx-local-draft', owner);
  wrapper.contentEditable = 'false';
  const field = document.createElement('textarea');
  field.value = 'Local caption';
  wrapper.append(field);
  return { wrapper, field };
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const portal of [false, true])
    test(`local draft focus hides only its owning session ${clientIds}/portal=${portal}`, async () => {
      const pair = await createTwoEditorFixture({ document: bytes, clientIds });
      const [alice, bob] = pair.editors;
      const host = alice!.surface!.hostElement();
      const { wrapper, field } = draft(
        host.ownerDocument,
        portal ? pair.rooms[0]!.session.sessionId : ''
      );
      const toolbar = host.ownerDocument.createElement('input');
      host.ownerDocument.body.append(toolbar);
      (portal ? host.ownerDocument.body : host).append(wrapper);
      const alicePresence = () =>
        pair.rooms[1]!.session.remoteSelections().find((value) => value.actorId === 'alice');
      const bobPresence = () =>
        pair.rooms[0]!.session.remoteSelections().find((value) => value.actorId === 'bob');
      try {
        const [first, second] = alice!.surface!.session.paragraphIds();
        alice!.surface!.setSelection({
          anchor: { paragraphId: first!, offset: 2 },
          head: { paragraphId: first!, offset: 2 },
        });
        bob!.surface!.setSelection({
          anchor: { paragraphId: second!, offset: 3 },
          head: { paragraphId: second!, offset: 3 },
        });
        await pair.settle();
        const shown = alicePresence();
        expect(shown).toBeDefined();
        expect(bobPresence()).toBeDefined();
        const painted = () =>
          pair.containers[1]!.querySelector<HTMLElement>('.docx-remote-caret-label');
        expect(painted()?.textContent).toBe('Alice');
        const geometry = { left: painted()!.style.left, top: painted()!.style.top };
        const before = await pair.savedDigests();
        const shared = pair.documents.map((doc) => Array.from(Y.encodeStateAsUpdate(doc)));
        const selection = alice!.surface!.state().selection;
        field.focus();
        field.setSelectionRange(1, 4);
        field.value = 'Uncommitted caption\nsecond line';
        field.dispatchEvent(new Event('input', { bubbles: true }));
        await pair.settle();
        expect(host.ownerDocument.activeElement).toBe(field);
        expect(alicePresence()).toBeUndefined();
        expect(bobPresence()).toBeDefined();
        expect(alice!.surface!.state().selection).toEqual(selection);
        expect(painted()).toBeNull();
        expect(await pair.savedDigests()).toEqual(before);
        expect(pair.documents.map((doc) => Array.from(Y.encodeStateAsUpdate(doc)))).toEqual(shared);
        alice!.focus();
        await pair.settle();
        expect(alicePresence()?.head).toEqual(shown!.head);
        expect(painted()?.textContent).toBe('Alice');
        expect({ left: painted()!.style.left, top: painted()!.style.top }).toEqual(geometry);
        toolbar.focus();
        await pair.settle();
        expect(alicePresence()?.head).toEqual(shown!.head);
        pair.pause();
        field.focus();
        await pair.settle();
        expect(alicePresence()?.head).toEqual(shown!.head);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        expect(alicePresence()).toBeUndefined();
        alice!.focus();
        await pair.settle();
        expect(alicePresence()?.head).toEqual(shown!.head);
        field.focus();
        await pair.settle();
        expect(alicePresence()).toBeUndefined();
        field.blur();
        await pair.settle();
        expect(alicePresence()?.head).toEqual(shown!.head);
        field.focus();
        await pair.settle();
        wrapper.remove();
        await pair.settle();
        expect(alicePresence()?.head).toEqual(shown!.head);
        expect(await pair.savedDigests()).toEqual(before);
      } finally {
        wrapper.remove();
        toolbar.remove();
        pair.destroy();
      }
    });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`portal draft ownership follows the current host window and stops on destroy ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({ document: bytes, clientIds });
    const alice = pair.editors[0]!;
    const host = alice.surface!.hostElement();
    const ownerDocument = host.ownerDocument;
    const portal = draft(ownerDocument, pair.rooms[0]!.session.sessionId);
    ownerDocument.body.append(portal.wrapper);
    const frame = ownerDocument.createElement('iframe');
    ownerDocument.body.append(frame);
    const targetDocument = frame.contentDocument!;
    const target = targetDocument.createElement('div');
    targetDocument.body.append(target);
    const nextPortal = draft(targetDocument, pair.rooms[0]!.session.sessionId);
    targetDocument.body.append(nextPortal.wrapper);
    const presence = () =>
      pair.rooms[1]!.session.remoteSelections().find((value) => value.actorId === 'alice');
    try {
      const paragraphId = alice.surface!.session.paragraphIds()[0]!;
      alice.surface!.setSelection({
        anchor: { paragraphId, offset: 2 },
        head: { paragraphId, offset: 2 },
      });
      await pair.settle();
      expect(presence()).toBeDefined();
      portal.field.focus();
      await pair.settle();
      expect(presence()).toBeUndefined();
      const owner = alice.surface;
      alice.attach(target, { focus: false });
      expect(alice.surface).toBe(owner);
      await pair.settle();
      expect(presence()).toBeDefined();
      expect(portal.wrapper.isConnected).toBe(true);
      portal.field.focus();
      await pair.settle();
      expect(presence()).toBeDefined();
      nextPortal.field.focus();
      await pair.settle();
      expect(presence()).toBeUndefined();
      portal.field.focus();
      await pair.settle();
      expect(presence()).toBeUndefined();
      const peer = pair.editors[1]!.surface!;
      expect(
        peer.applyAutomationOps(
          () => [{ op: 'deleteBlock', blockId: peer.session.paragraphIds()[0]! }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).committed
      ).toBe(true);
      await pair.settle();
      expect(targetDocument.activeElement).toBe(nextPortal.field);
      expect(presence()).toBeUndefined();
      nextPortal.wrapper.remove();
      alice.focus();
      await pair.settle();
      expect(presence()).toBeDefined();
      targetDocument.body.append(nextPortal.wrapper);
      nextPortal.field.focus();
      await pair.settle();
      expect(presence()).toBeUndefined();
      alice.destroy();
      portal.field.focus();
      portal.field.blur();
      await pair.settle();
      expect(presence()).toBeUndefined();
    } finally {
      portal.wrapper.remove();
      nextPortal.wrapper.remove();
      frame.remove();
      pair.destroy();
    }
  });
