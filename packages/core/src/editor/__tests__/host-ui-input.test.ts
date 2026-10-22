// HOST UI INPUT NEVER REACHES CAPTURE-PHASE DOCUMENT GESTURES.
//
// The pages layer reads some key presses in its capture phase, before a host can stop them on
// its own container. A selected picture used to take Backspace from a widget textarea painted
// inside the pages: the press deleted the picture, and the textarea lost its own edit. The
// painted caret also stayed on screen while the textarea held focus, and the pages kept their
// transparent caret color, which the textarea inherited.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { validateRasterHeader, type ImageDecodePort } from '../../store/package/image-resources.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { timedPress } from '../surface-timing.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const PNG_1X1 = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decoder: ImageDecodePort = Object.freeze({
  async decode(bytes: Uint8Array, mime: string) {
    const header = validateRasterHeader(bytes, mime);
    if (!header) throw new Error('invalid image');
    return Object.freeze({
      pixelWidth: header.pixelWidth,
      pixelHeight: header.pixelHeight,
      dpiX: 96,
      dpiY: 96,
    });
  },
});

function pictureDocument(): Uint8Array {
  const drawing =
    '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/>' +
    '<wp:docPr id="1" name="green"/><wp:cNvGraphicFramePr/>' +
    `<a:graphic xmlns:a="${A}"><a:graphicData uri="${PIC}">` +
    `<pic:pic xmlns:pic="${PIC}"><pic:nvPicPr><pic:cNvPr id="1" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rIdImage"><a:stretch><a:fillRect/></a:stretch></a:blip></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline>';
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:r="${R}"><w:body>` +
        `<w:p><w:r><w:t>before</w:t></w:r><w:r><w:drawing>${drawing}</w:drawing></w:r><w:r><w:t>after</w:t></w:r></w:p>` +
        '</w:body></w:document>'
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rIdImage" Type="${R}/image" Target="media/image1.png"/></Relationships>`
    ),
    'word/media/image1.png': PNG_1X1,
  });
}

function mounted() {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: pictureDocument(),
    imageDecodePort: decoder,
  });
  const surface = editor.surface!;
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  const content = pages.querySelector<HTMLElement>('.docx-page-content')!;
  const ui = document.createElement('div');
  ui.setAttribute('data-docx-host-ui', '');
  const field = document.createElement('textarea');
  ui.append(field);
  content.append(ui);
  // Hosts stop bubbling at their container, like the engine's own control menus.
  for (const type of ['keydown', 'beforeinput', 'pointerdown'])
    ui.addEventListener(type, (event) => event.stopPropagation());
  return { editor, surface, container, pages, ui, field };
}

/** Select the picture the way a press on it does: model position plus pointer intent. */
function selectPicture(editor: DocxEditorInstance, pages: HTMLElement): string {
  const surface = editor.surface!;
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId, offset: 6 },
    head: { paragraphId, offset: 6 },
  });
  const id = editor.snapshot().image!.id;
  // The capture-phase object reader arms pointer intent. No text hit test runs afterward.
  const painted = document.createElement('span');
  painted.setAttribute('data-drawing-node-id', id);
  painted.addEventListener('pointerdown', (event) => event.stopPropagation());
  pages.querySelector('.docx-page-content')!.append(painted);
  painted.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
  painted.remove();
  expect(editor.snapshot().image?.id).toBe(id);
  return id;
}

const backspace = () =>
  new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true, cancelable: true });

describe('host UI input stays outside document gestures', () => {
  test('Backspace in host UI leaves a selected picture and the text field press intact', () => {
    const { editor, surface, container, pages, ui, field } = mounted();
    try {
      const id = selectPicture(editor, pages);
      const revision = surface.session.packageRevision();
      const press = backspace();
      field.dispatchEvent(press);
      expect(press.defaultPrevented).toBe(false);
      expect(surface.session.packageRevision()).toBe(revision);
      expect(editor.snapshot().image?.id).toBe(id);
      // Control: without the marker, the capture-phase reader takes the same press.
      ui.removeAttribute('data-docx-host-ui');
      selectPicture(editor, pages);
      const leaked = backspace();
      field.dispatchEvent(leaked);
      expect(leaked.defaultPrevented).toBe(true);
      expect(surface.session.packageRevision()).toBeGreaterThan(revision);
    } finally {
      editor.destroy();
      container.remove();
    }
  });

  test('host UI presses are not timed as editor input', () => {
    const { editor, container, pages, field } = mounted();
    try {
      const inside = new KeyboardEvent('keydown', { key: 'a', bubbles: true });
      field.dispatchEvent(inside);
      expect(timedPress(inside)).toBe(false);
      const outside = new KeyboardEvent('keydown', { key: 'a', bubbles: true });
      pages.dispatchEvent(outside);
      expect(timedPress(outside)).toBe(true);
    } finally {
      editor.destroy();
      container.remove();
    }
  });

  test('the painted caret yields while host UI holds focus', () => {
    const { editor, surface, container, pages, field } = mounted();
    try {
      const paragraphId = surface.session.paragraphIds()[0]!;
      pages.focus();
      surface.setSelection({
        anchor: { paragraphId, offset: 2 },
        head: { paragraphId, offset: 2 },
      });
      expect(pages.querySelector('[data-docx-caret]')).not.toBeNull();
      field.focus();
      expect(document.activeElement).toBe(field);
      expect(pages.querySelector('[data-docx-caret]')).toBeNull();
      expect(pages.style.getPropertyValue('caret-color')).toBe('');
      pages.focus();
      surface.setSelection({
        anchor: { paragraphId, offset: 3 },
        head: { paragraphId, offset: 3 },
      });
      expect(pages.querySelector('[data-docx-caret]')).not.toBeNull();
    } finally {
      editor.destroy();
      container.remove();
    }
  });
});
