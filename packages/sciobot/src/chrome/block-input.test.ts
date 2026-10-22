import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  directParagraphProperties,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { withBlockTypeStyles } from './block-styles';
import { installBlockInputRules } from './block-input';
import {
  findNode,
  bodyStoryRoot,
  storyParagraphs,
  canonicalOoxmlFingerprint,
} from '@docx-editor.dev/core/store';
import { isHorizontalRuleParagraph } from '../native/horizontal-rule';
import reference from '../__tests__/fixtures/native-commands.json';

function opened(text: string, style?: string) {
  const loaded = readOoxmlPackage(
    docxFixture(
      `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const prepared = withBlockTypeStyles(loaded.package);
  if (!prepared) throw new Error('Missing prepared block definitions.');
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(prepared),
  });
  const stop = installBlockInputRules(editor, { documentId: 'block-input', epoch: 'mount' });
  const surface = editor.surface!,
    id = surface.session.paragraphIds()[0]!;
  const point = { paragraphId: id, offset: text.length };
  surface.setSelection({ anchor: point, head: point });
  const target = surface.hostElement().querySelector<HTMLElement>('[contenteditable="true"]');
  if (!target) throw new Error('Missing native input surface.');
  return {
    editor,
    surface,
    id,
    target,
    destroy() {
      stop();
      editor.destroy();
    },
  };
}
for (const [id, prefix, style] of [
  ['input.heading1', '#', 'Heading1'],
  ['input.heading2', '##', 'Heading2'],
  ['input.heading3', '###', 'Heading3'],
  ['input.quote', '>', 'Quote'],
] as const)
  test(`${id} consumes the captured prefix and retains the paragraph identity`, () => {
    const fixture = reference.cases.find((entry) => entry.id === id)!;
    expect(fixture.arguments).toEqual({ text: (prefix === '>' ? '>' : prefix) + ' ' });
    const current = opened(prefix);
    try {
      const event = new InputEvent('beforeinput', {
        bubbles: true,
        cancelable: true,
        inputType: 'insertText',
        data: ' ',
      });
      current.target.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
      expect(current.surface.session.paragraphIds()).toEqual([current.id]);
      expect(paragraphTextOf(current.surface.session.part(), current.id)).toBe('');
      expect(
        directParagraphProperties(current.surface.session.part(), current.id).find(
          (p) => p.localName === 'pStyle'
        )?.attributes?.val
      ).toBe(style);
      expect(current.surface.state().selection.head).toEqual({
        paragraphId: current.id,
        offset: 0,
      });
      expect(current.editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(paragraphTextOf(current.surface.session.part(), current.id)).toBe(
        prefix === '>' ? '>' : prefix
      );
      expect(current.editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(paragraphTextOf(current.surface.session.part(), current.id)).toBe('');
    } finally {
      current.destroy();
    }
  });
for (const style of ['Heading1', 'Heading2', 'Heading3'])
  test(`Enter resets empty ${style} in place with one undo`, () => {
    const current = opened('', style);
    try {
      const event = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
      current.target.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
      expect(current.surface.session.paragraphIds()).toEqual([current.id]);
      expect(
        directParagraphProperties(current.surface.session.part(), current.id).find(
          (p) => p.localName === 'pStyle'
        )?.attributes?.val
      ).toBe('Normal');
      expect(current.editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(
        directParagraphProperties(current.surface.session.part(), current.id).find(
          (p) => p.localName === 'pStyle'
        )?.attributes?.val
      ).toBe(style);
    } finally {
      current.destroy();
    }
  });

for (const [prefix, trigger] of [
  ['--', '-'],
  ['___', ' '],
] as const)
  test(`rule prefix ${prefix} allocates a native rule and follower atomically`, async () => {
    const current = opened(prefix);
    try {
      const event = new InputEvent('beforeinput', {
        bubbles: true,
        cancelable: true,
        inputType: 'insertText',
        data: trigger,
      });
      current.target.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
      const ids = current.surface.session.paragraphIds();
      expect(ids).toHaveLength(2);
      expect(ids[1]).toBe(current.id);
      expect(isHorizontalRuleParagraph(findNode(current.surface.session.part(), ids[0]!)!)).toBe(
        true
      );
      expect(paragraphTextOf(current.surface.session.part(), ids[1]!)).toBe('');
      expect(current.surface.state().selection.head).toEqual({ paragraphId: ids[1]!, offset: 0 });
      const loaded = readOoxmlPackage(new Uint8Array(await current.editor.save()));
      if (!loaded.ok) throw new Error(loaded.reason);
      const saved = loaded.package.parts.get(loaded.package.mainDocumentPart)!;
      const savedRule = storyParagraphs(bodyStoryRoot(saved)!)[0]!;
      expect(isHorizontalRuleParagraph(savedRule)).toBe(true);
      expect(canonicalOoxmlFingerprint(savedRule)).toBe(
        canonicalOoxmlFingerprint(findNode(current.surface.session.part(), ids[0]!)!)
      );
      expect(current.editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(current.surface.session.paragraphIds()).toEqual([current.id]);
      expect(paragraphTextOf(current.surface.session.part(), current.id)).toBe(prefix);
      expect(current.editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(isHorizontalRuleParagraph(findNode(current.surface.session.part(), ids[0]!)!)).toBe(
        true
      );
    } finally {
      current.destroy();
    }
  });
