// A RICH PASTE CAN EXTEND AN EXPLICIT HISTORY GROUP.
//
// A fragment paste lands as a package undo unit. That unit used to ignore the history group of
// the commit in flight, so an adapter that repaired the pasted content in a second write left
// two undo steps. One undo then exposed the unrepaired paste.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { beginHistoryGroupFrames } from '../editor-history-group-frames.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const editors: DocxEditorInstance[] = [];
function open(text: string): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(`<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`),
  });
  editors.push(editor);
  return editor;
}
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

function richFlavours() {
  const source = open('Pasted');
  source.surface!.selectAll();
  const flavours = source.surface!.copyFlavours();
  expect(flavours.html).toContain('data-docx-fragment');
  return flavours;
}

/** Paste at the end of `AB`, then write `!` at the start, as two native writes. */
function pasteAndRepair(target: DocxEditorInstance, run: <T>(write: () => T) => T) {
  const surface = target.surface!;
  const id = surface.session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 2 },
    head: { paragraphId: id, offset: 2 },
  });
  const flavours = richFlavours();
  expect(run(() => surface.pasteRich(flavours.text, flavours.html))).toBe(true);
  const repaired = run(() =>
    surface.applyAutomationOps(() => [{ op: 'insertText', paragraphId: id, offset: 0, text: '!' }])
  );
  expect(repaired.rejected).toBe(false);
  expect(surface.session.bodyText()).toBe('!ABPasted\n');
}

test('a rich paste and a follow-up write share one undo step inside history group frames', () => {
  const target = open('AB');
  const group = target.beginHistoryGroup();
  const frames = beginHistoryGroupFrames(target, group)!;
  expect(frames).not.toBeNull();
  pasteAndRepair(target, (write) => frames.run(write));
  expect(frames.finish()).toEqual({ kind: 'started' });
  expect(group.lastOutcome).toEqual({ kind: 'started' });
  group.end();
  const surface = target.surface!;
  surface.undo();
  expect(surface.session.bodyText()).toBe('AB');
  surface.redo();
  expect(surface.session.bodyText()).toBe('!ABPasted\n');
  // A closed handle cannot start new frames.
  expect(beginHistoryGroupFrames(target, group)).toBeNull();
});

test('without frames, the same writes remain two undo steps', () => {
  const target = open('AB');
  pasteAndRepair(target, (write) => write());
  const surface = target.surface!;
  surface.undo();
  expect(surface.session.bodyText()).toBe('ABPasted\n');
  surface.undo();
  expect(surface.session.bodyText()).toBe('AB');
});

test('frames refuse a handle from another editor', () => {
  const first = open('A'),
    second = open('B');
  const group = first.beginHistoryGroup();
  expect(beginHistoryGroupFrames(second, group)).toBeNull();
  group.end();
});
