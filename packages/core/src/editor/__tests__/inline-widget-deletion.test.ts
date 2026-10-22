import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { setSurfaceInlineWidgetMetrics } from '../surface-inline-widget-metrics.ts';
import { absorbAtomicControls } from '../surface-atomic-controls.ts';
import { docx } from '../../automation/__tests__/support/protocol.ts';
import {
  canonicalOoxmlFingerprint,
  contentControlsIn,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
} from '../../store/index.ts';

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
function mount(lock = '') {
  const body = `<w:p>${run('A')}<w:sdt><w:sdtPr><w:tag w:val="chip"/>${lock ? `<w:lock w:val="${lock}"/>` : ''}</w:sdtPr><w:sdtContent>${run('source')}</w:sdtContent></w:sdt>${run('B')}</w:p>`;
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docx(body),
    author: 'Ada',
  });
  const surface = editor.surface!,
    id = surface.session.paragraphIds()[0]!;
  setSurfaceInlineWidgetMetrics(surface, {
    measure: () => ({ width: 72, ascent: 18, descent: 2 }),
    revision: '1',
  });
  return {
    editor,
    surface,
    id,
    select: (anchor: number, head = anchor) =>
      surface.setSelection({
        anchor: { paragraphId: id, offset: anchor },
        head: { paragraphId: id, offset: head },
      }),
    text: () => paragraphTextOf(surface.session.part(), id),
    fingerprint: () => canonicalOoxmlFingerprint(surface.session.part().root),
  };
}
for (const direction of ['backward', 'forward'] as const)
  test(`one native ${direction} delete removes the measured control with exact history and Word roundtrip`, async () => {
    const m = mount();
    try {
      const before = m.fingerprint();
      m.select(direction === 'backward' ? 7 : 1);
      if (direction === 'backward') m.surface.deleteBackward();
      else m.surface.deleteForward();
      expect(m.text()).toBe('AB');
      expect([...contentControlsIn(m.surface.session.part().root)]).toHaveLength(0);
      expect(m.surface.state().selection).toMatchObject({ head: { offset: 1 } });
      const removed = m.fingerprint();
      expect(m.editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(m.fingerprint()).toBe(before);
      expect(m.editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(m.fingerprint()).toBe(removed);
      const saved = readOoxmlPackage(new Uint8Array(await m.editor.save()));
      if (!saved.ok) throw new Error(saved.reason);
      expect(
        canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!.root)
      ).toBe(removed);
    } finally {
      m.editor.destroy();
    }
  });

for (const [from, to, expected] of [
  [1, 7, 'AB'],
  [0, 7, 'B'],
  [1, 8, 'A'],
] as const)
  test(`native selected deletion removes the whole widget and surrounding range ${from}/${to}`, () => {
    const m = mount();
    try {
      const before = m.fingerprint();
      m.select(from, to);
      expect(m.surface.deleteSelection()).toBe(true);
      expect(m.text()).toBe(expected);
      expect([...contentControlsIn(m.surface.session.part().root)]).toHaveLength(0);
      expect(m.editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(m.fingerprint()).toBe(before);
      const partial = absorbAtomicControls(m.surface.layout(), {
        anchor: { paragraphId: m.id, offset: 2 },
        head: { paragraphId: m.id, offset: 5 },
      });
      expect(partial).toEqual({
        anchor: { paragraphId: m.id, offset: 1 },
        head: { paragraphId: m.id, offset: 7 },
      });
    } finally {
      m.editor.destroy();
    }
  });

test('clearing host metrics restores native character deletion without changing the control tag', () => {
  const m = mount();
  try {
    setSurfaceInlineWidgetMetrics(m.surface, null);
    m.select(7);
    m.surface.deleteBackward();
    expect(m.text()).toBe('AsourcB');
    expect([...contentControlsIn(m.surface.session.part().root)]).toHaveLength(1);
  } finally {
    m.editor.destroy();
  }
});

for (const lock of ['contentLocked', 'sdtLocked', 'sdtContentLocked'])
  test(`measured atoms retain native ${lock} removal rules`, () => {
    const m = mount(lock);
    try {
      const before = m.fingerprint();
      m.select(7);
      m.surface.deleteBackward();
      if (lock !== 'sdtLocked') {
        expect(m.surface.state().selection).toMatchObject({
          anchor: { offset: 7 },
          head: { offset: 1 },
        });
        expect(m.fingerprint()).toBe(before);
        m.surface.deleteBackward();
      }
      if (lock === 'contentLocked') expect(m.text()).toBe('AB');
      else expect(m.fingerprint()).toBe(before);
    } finally {
      m.editor.destroy();
    }
  });

test('suggested deletion retains the native SDT wrapper and tracks its complete fallback with exact undo', () => {
  const m = mount();
  try {
    const before = m.fingerprint();
    m.surface.setEditingMode('suggest');
    m.select(7);
    m.surface.deleteBackward();
    const xml = serializeOoxmlPart(m.surface.session.part());
    // Native unlocked SDTs retain their wrapper and track the complete fallback inside it.
    expect(xml).toMatch(/<w:sdtContent><w:del\b[^>]*w:author="Ada"[^>]*><w:r>/);
    expect(xml).toContain('<w:delText>source</w:delText>');
    expect(m.text()).toBe('AsourceB');
    expect(m.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(m.fingerprint()).toBe(before);
  } finally {
    m.editor.destroy();
  }
});
