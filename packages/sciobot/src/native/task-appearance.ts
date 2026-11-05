import type {
  ParagraphAppearanceContribution,
  ParagraphAppearanceContext,
} from '@docx-editor.dev/core/editor';
import { checklistTag } from '../registry/adapter-names';

const registrations = new Map<string, ParagraphAppearanceContribution>();

/** Capture the startup registry once; the registered resolver reads only immutable native facts. */
export function taskAppearanceContribution(): ParagraphAppearanceContribution {
  const tag = checklistTag();
  const existing = registrations.get(tag);
  if (existing) return existing;
  const contribution: ParagraphAppearanceContribution = Object.freeze({
    id: `task-appearance:${tag}`,
    resolve(context: ParagraphAppearanceContext) {
      const markers = context.controls.filter(
        (control) => control.tag === tag && control.controlType === 'checkbox'
      );
      return markers.length === 1 && markers[0]!.checked
        ? { strike: true, color: { val: '606060', themeColor: 'text1', themeTint: '60' } }
        : null;
    },
  });
  if (registrations.size >= 32) registrations.delete(registrations.keys().next().value!);
  registrations.set(tag, contribution);
  return contribution;
}
