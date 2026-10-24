import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor';
import { canonicalOoxmlFingerprint, readOoxmlPackage, serializeOoxmlPart } from '../../store/index';

test('caret character style uses the explicit operation field and independent typing history', async () => {
  const editor = createDocxEditor({ container: document.createElement('div'), document: 'blank' });
  try {
    const s = editor.surface!;
    const before = canonicalOoxmlFingerprint(s.session.part());
    expect(s.setRunProperty('rStyle', { val: 'ImportedCode' })).toBeNull();
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect<unknown>(s.state().pendingFormat).toEqual([
      { localName: 'rStyle', attributes: { val: 'ImportedCode' } },
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(s.state().pendingFormat).toBeNull();
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    s.insertPlainText('Code');
    s.flushPendingInput();
    expect(serializeOoxmlPart(s.session.part())).toContain('w:val="ImportedCode"');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(serializeOoxmlPart(saved.package.parts.get(saved.package.mainDocumentPart)!)).toContain(
      'w:val="ImportedCode"'
    );
  } finally {
    editor.destroy();
  }
});

test('caret character style merges with other armed properties, clears explicitly, and refuses invalid identities', () => {
  const editor = createDocxEditor({ container: document.createElement('div'), document: 'blank' });
  try {
    const s = editor.surface!;
    expect(s.setRunProperty('rStyle', { val: 'Code' })).toBeNull();
    expect(s.setRunProperty('b')).toBeNull();
    s.insertPlainText('Code');
    s.flushPendingInput();
    expect(serializeOoxmlPart(s.session.part())).toContain('w:val="Code"');
    expect(s.formatting().bold).toBe(true);
    expect(s.setRunProperty('rStyle', { val: '' })).toBeNull();
    s.insertPlainText('Plain');
    s.flushPendingInput();
    const id = s.session.paragraphIds()[0]!;
    s.setSelection({
      anchor: { paragraphId: id, offset: 4 },
      head: { paragraphId: id, offset: 9 },
    });
    expect(s.setRunProperty('rStyle', { val: 'Range' })).toBe(
      'Use an explicit characterStyle operation for a range.'
    );
    const before = canonicalOoxmlFingerprint(s.session.part());
    s.setSelection({
      anchor: { paragraphId: id, offset: 9 },
      head: { paragraphId: id, offset: 9 },
    });
    expect(s.setRunProperty('rStyle', { val: 'x'.repeat(513) })).not.toBeNull();
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    const xml = serializeOoxmlPart(s.session.part());
    expect(xml).toContain('Plain');
    expect(xml).not.toContain('w:val="Range"');
    const plainRun = xml.slice(xml.indexOf('Code</w:t>') + 11);
    expect(plainRun).not.toContain('w:val="Code"');
  } finally {
    editor.destroy();
  }
});

test('Enter retains an armed character style without passing it through paragraph mark properties', () => {
  const editor = createDocxEditor({ container: document.createElement('div'), document: 'blank' });
  try {
    const s = editor.surface!;
    s.insertPlainText('Before');
    s.flushPendingInput();
    expect(s.setRunProperty('rStyle', { val: 'Code' })).toBeNull();
    s.splitParagraph();
    expect(s.session.paragraphIds()).toHaveLength(2);
    s.insertPlainText('After');
    s.flushPendingInput();
    const xml = serializeOoxmlPart(s.session.part());
    expect(xml).toContain('After');
    expect(xml).toContain('w:val="Code"');
    expect(xml).not.toContain('<w:pPr><w:rPr><w:rStyle');
    expect(s.state().lastRejection).toBeNull();
  } finally {
    editor.destroy();
  }
});
