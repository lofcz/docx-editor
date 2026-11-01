import { expect, test } from 'bun:test';
import { act } from 'react';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { nativeSlotDocument } from '../../test/support/react-block-document';
import { mountReactBlockHost, type ReactBlockRegistry, type ReactBlockViewProps } from './host';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
const settle = () => new Promise((resolve) => setTimeout(resolve, 30));

function Frame({ children }: ReactBlockViewProps) {
  return <div style={{ display: 'contents' }}>{children}</div>;
}
/** 20 px left and 8 px right: 15 pt and 6 pt. */
const REGISTRY: ReactBlockRegistry = new Map([
  ['native_container', { view: Frame, inset: { left: 20, right: 8 } }],
]);

function idOf(editor: DocxEditorInstance, text: string) {
  const session = editor.surface!.session;
  const id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
  if (!id) throw new Error(`Missing paragraph ${text}`);
  return id;
}
function xOf(editor: DocxEditorInstance, text: string) {
  editor.surface!.flushPendingInput();
  return caretAt(editor.surface!.layout(), { paragraphId: idOf(editor, text), offset: 0 })!.x;
}

test('a registered inset pads container content, adds up through nesting, and edits nothing', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: nativeSlotDocument() });
  try {
    const texts = ['Intro', 'Outer start', 'Inner text', 'Cell text', 'Leaf text', 'Tail'];
    const before = Object.fromEntries(texts.map((text) => [text, xOf(editor, text)]));
    const fingerprint = canonicalOoxmlFingerprint(editor.surface!.session.part());
    let host: ReturnType<typeof mountReactBlockHost> | undefined;
    await act(async () => {
      host = mountReactBlockHost(editor, REGISTRY);
      await settle();
    });
    const after = Object.fromEntries(texts.map((text) => [text, xOf(editor, text)]));
    expect(after['Intro']).toBe(before['Intro']!);
    expect(after['Tail']).toBe(before['Tail']!);
    expect(after['Outer start']).toBe(before['Outer start']! + 15);
    expect(after['Inner text']).toBe(before['Inner text']! + 30);
    // Content in a table cell follows its inset table.
    expect(after['Cell text']).toBe(before['Cell text']! + 15);
    // `only-child` holds `leaf`: three nested containers.
    expect(after['Leaf text']).toBe(before['Leaf text']! + 45);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
    expect(editor.surface!.state().canUndo).toBe(false);
    // A new paragraph in a container is inset as soon as it exists.
    const start = idOf(editor, 'Outer start');
    await act(async () => {
      editor.surface!.setSelection({
        anchor: { paragraphId: start, offset: 11 },
        head: { paragraphId: start, offset: 11 },
      });
      editor.surface!.splitParagraph();
      editor.surface!.insertPlainText('Typed');
      await settle();
    });
    expect(xOf(editor, 'Typed')).toBe(before['Outer start']! + 15);
    // Disposal removes every inset.
    await act(async () => {
      host!.dispose();
      await settle();
    });
    expect(xOf(editor, 'Inner text')).toBe(before['Inner text']!);
  } finally {
    editor.destroy();
    container.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});

for (const receiver of [0, 1] as const)
  test(`two sessions inset a peer's new container paragraph alike; receiver ${receiver}`, async () => {
    const previous = environment.IS_REACT_ACT_ENVIRONMENT;
    environment.IS_REACT_ACT_ENVIRONMENT = true;
    const pair = await createTwoEditorFixture({
      document: nativeSlotDocument(),
      clientIds: [73, 41],
    });
    const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
    try {
      await act(async () => {
        pair.editors.forEach((editor) => hosts.push(mountReactBlockHost(editor, REGISTRY)));
        await settle();
      });
      const [alice, bob] = pair.editors;
      const inner = xOf(alice!, 'Inner text');
      pair.pause();
      await act(async () => {
        const id = idOf(bob!, 'Inner text');
        bob!.surface!.setSelection({
          anchor: { paragraphId: id, offset: 10 },
          head: { paragraphId: id, offset: 10 },
        });
        bob!.surface!.splitParagraph();
        bob!.surface!.insertPlainText('Bob line');
        const outer = idOf(alice!, 'Outer middle');
        alice!.surface!.setSelection({
          anchor: { paragraphId: outer, offset: 0 },
          head: { paragraphId: outer, offset: 0 },
        });
        alice!.surface!.insertPlainText('Alice ');
        pair.reconnect(receiver);
        await pair.settle();
        await settle();
      });
      for (const editor of pair.editors) {
        expect(xOf(editor, 'Bob line')).toBe(inner);
        expect(xOf(editor, 'Alice Outer middle')).toBe(inner - 15);
      }
      // Undo on the author's session removes the line from both; insets stay layout-only.
      await act(async () => {
        bob!.surface!.undo();
        bob!.surface!.flushPendingInput();
        await pair.settle();
        await settle();
      });
      for (const editor of pair.editors)
        expect(
          editor
            .surface!.session.paragraphIds()
            .some((id) => paragraphTextOf(editor.surface!.session.part(), id) === 'Bob line')
        ).toBe(false);
      expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(
        canonicalOoxmlFingerprint(bob!.surface!.session.part())
      );
    } finally {
      await act(async () => {
        hosts.forEach((host) => host.dispose());
      });
      pair.destroy();
      environment.IS_REACT_ACT_ENVIRONMENT = previous;
    }
  });

test('top and bottom room go to the first and last paragraph of a container', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: nativeSlotDocument() });
  const yOf = (text: string) => {
    editor.surface!.flushPendingInput();
    return caretAt(editor.surface!.layout(), { paragraphId: idOf(editor, text), offset: 0 })!.y;
  };
  try {
    const gap = (from: string, to: string) => yOf(to) - yOf(from);
    const before = { open: gap('Intro', 'Outer start'), middle: gap('Outer start', 'Inner text') };
    let host: ReturnType<typeof mountReactBlockHost> | undefined;
    await act(async () => {
      host = mountReactBlockHost(
        editor,
        new Map([['native_container', { view: Frame, inset: { top: 80 } }]])
      );
      await settle();
    });
    // 80 px is 60 pt: more than any authored spacing here, so it decides the gap.
    expect(gap('Intro', 'Outer start')).toBeGreaterThan(before.open + 40);
    // `Inner text` opens the `inner` container, so it gets room too; `Outer middle` does not.
    expect(gap('Outer start', 'Inner text')).toBeGreaterThan(before.middle + 40);
    expect(xOf(editor, 'Inner text')).toBe(xOf(editor, 'Intro'));
    await act(async () => {
      host!.dispose();
      await settle();
    });
    expect(gap('Intro', 'Outer start')).toBeCloseTo(before.open, 6);
  } finally {
    editor.destroy();
    container.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});
