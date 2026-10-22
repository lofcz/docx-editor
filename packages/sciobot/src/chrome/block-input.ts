import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { directParagraphProperties, paragraphTextOf } from '@docx-editor.dev/core/store';
import { setBlockType, type BlockCommandContext, type TextBlockType } from './block-type';
import { applyHorizontalRuleInput } from './rule-input';

/** Private application rules run before the shared native keymap. Disposal removes every listener. */
export function installBlockInputRules(editor: DocxEditorInstance, context: BlockCommandContext) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const host = surface.hostElement();
  const inSurface = (event: Event) => {
    const target = event.target as Element | null;
    return (
      !!target?.closest?.('[contenteditable="true"]') && !target.closest('[data-docx-block-ui]')
    );
  };
  const selected = () => {
    surface.flushPendingInput();
    const { anchor, head } = surface.state().selection;
    if (anchor.paragraphId !== head.paragraphId || anchor.offset !== head.offset) return null;
    const part = surface.session.part();
    if (part.name !== surface.session.currentPackage().mainDocumentPart) return null;
    return {
      point: head,
      text: paragraphTextOf(part, head.paragraphId),
      style: directParagraphProperties(part, head.paragraphId).find(
        (value) => value.localName === 'pStyle'
      )?.attributes?.val,
    };
  };
  const stop = (event: Event) => {
    event.preventDefault();
    event.stopImmediatePropagation();
  };
  const beforeInput = (event: InputEvent) => {
    if (!inSurface(event) || event.isComposing || event.inputType !== 'insertText') return;
    const current = selected();
    if (
      !current ||
      (current.style && current.style !== 'Normal') ||
      current.point.offset !== current.text?.length
    )
      return;
    if (
      (event.data === '-' && current.text === '--') ||
      (event.data === ' ' && current.text === '___')
    ) {
      if (applyHorizontalRuleInput(editor, current.text as '--' | '___')) stop(event);
      return;
    }
    if (event.data !== ' ') return;
    const types: Readonly<Record<string, TextBlockType>> = {
      '#': 'h1',
      '##': 'h2',
      '###': 'h3',
      '>': 'blockquote',
    };
    const type = current.text && types[current.text];
    if (!type) return;
    const result = setBlockType(editor, type, context, current.text!);
    if (!result.success) return;
    surface.setSelection({
      anchor: { ...current.point, offset: 0 },
      head: { ...current.point, offset: 0 },
    });
    stop(event);
  };
  const keydown = (event: KeyboardEvent) => {
    if (!inSurface(event) || event.isComposing) return;
    const accelerator = event.ctrlKey || event.metaKey;
    const heading = accelerator && event.altKey && !event.shiftKey && /^[1-3]$/.test(event.key);
    const quote = accelerator && event.shiftKey && !event.altKey && event.code === 'Period';
    if ((heading || quote) && !event.getModifierState('AltGraph')) {
      surface.flushPendingInput();
      const type: TextBlockType = heading ? (`h${event.key}` as TextBlockType) : 'blockquote';
      const style = directParagraphProperties(
        surface.session.part(),
        surface.state().selection.head.paragraphId
      ).find((value) => value.localName === 'pStyle')?.attributes?.val;
      const targetStyle = type === 'blockquote' ? 'Quote' : `Heading${event.key}`;
      if (setBlockType(editor, style === targetStyle ? 'p' : type, context).success) stop(event);
      return;
    }
    if (event.key !== 'Enter' || event.shiftKey || event.ctrlKey || event.metaKey || event.altKey)
      return;
    const current = selected();
    if (current?.text !== '' || !/^Heading[1-3]$/.test(current.style ?? '')) return;
    if (setBlockType(editor, 'p', context).success) stop(event);
  };
  host.addEventListener('beforeinput', beforeInput, true);
  host.addEventListener('keydown', keydown, true);
  return () => {
    host.removeEventListener('beforeinput', beforeInput, true);
    host.removeEventListener('keydown', keydown, true);
  };
}
