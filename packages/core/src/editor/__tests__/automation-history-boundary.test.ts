import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { blankDocumentBytes } from '../blank-document.ts';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  customXmlDataParts,
  customXmlNodes,
  withCustomXmlDataPart,
  withCustomXmlNode,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '../../store/index.ts';

const NS = 'urn:test:automation-history';
const mount = () =>
  createDocxEditor({ container: document.createElement('div'), document: blankDocumentBytes() });
function metadata(pkg: OoxmlPackage, id: string) {
  const created = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, NS, 'nodes');
  if (!created.part) throw new Error('Missing native metadata part');
  return withCustomXmlNode(created.pkg, created.part.partName, {
    id,
    label: id,
    data: JSON.stringify({ id }),
  });
}
function ids(pkg: OoxmlPackage) {
  return customXmlDataParts(pkg, pkg.mainDocumentPart)
    .filter((part) => part.namespaceUri === NS)
    .flatMap((part) => customXmlNodes(pkg, part.partName).map((node) => node.id))
    .sort();
}

test('explicit automation calls create independent undo units for body and package edits without timing delays', () => {
  const editor = mount();
  try {
    const surface = editor.surface!,
      session = surface.session,
      id = session.paragraphIds()[0]!,
      original = canonicalOoxmlFingerprint(session.part());
    const first = surface.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId: id, offset: 0, text: 'A' }],
      { kind: 'body' },
      [(pkg) => metadata(pkg, 'a')],
      false,
      true
    );
    const second = surface.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId: id, offset: 1, text: 'B' }],
      { kind: 'body' },
      [(pkg) => metadata(pkg, 'b')],
      false,
      true
    );
    expect([first.committed, second.committed]).toEqual([true, true]);
    expect(paragraphTextOf(session.part(), id)).toBe('AB');
    expect(ids(session.currentPackage())).toEqual(['a', 'b']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('A');
    expect(ids(session.currentPackage())).toEqual(['a']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(original);
    expect(ids(session.currentPackage())).toEqual([]);
    expect(session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('A');
    expect(ids(session.currentPackage())).toEqual(['a']);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('AB');
    expect(ids(session.currentPackage())).toEqual(['a', 'b']);
  } finally {
    editor.destroy();
  }
});

test('default automation batches retain the existing local history behavior', () => {
  const editor = mount();
  try {
    const surface = editor.surface!,
      id = surface.session.paragraphIds()[0]!;
    for (const [offset, text] of [
      [0, 'A'],
      [1, 'B'],
    ] as const)
      expect(
        surface.applyAutomationOps(() => [{ op: 'insertText', paragraphId: id, offset, text }])
          .committed
      ).toBe(true);
    expect(paragraphTextOf(surface.session.part(), id)).toBe('AB');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(surface.session.part(), id)).toBe('A');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(surface.session.part(), id)).toBe('');
    expect(surface.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('no-op, refused, and viewing calls preserve package bytes and add no undo units', () => {
  const editor = mount();
  try {
    const surface = editor.surface!,
      session = surface.session,
      bytes = writeOoxmlPackage(session.currentPackage());
    expect(
      surface.applyAutomationOps(() => [], { kind: 'body' }, undefined, false, true).committed
    ).toBe(false);
    expect(
      surface.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: 'missing', offset: 0, text: 'No' }],
        { kind: 'body' },
        [(pkg) => metadata(pkg, 'invalid')],
        false,
        true
      ).rejected
    ).toBe(true);
    surface.setEditingMode('view');
    let called = false;
    expect(
      surface.applyAutomationOps(
        () => {
          called = true;
          return [];
        },
        { kind: 'body' },
        [(pkg) => metadata(pkg, 'view')],
        false,
        true
      ).rejected
    ).toBe(true);
    expect(called).toBe(false);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect([session.canUndo(), session.canRedo()]).toEqual([false, false]);
  } finally {
    editor.destroy();
  }
});
