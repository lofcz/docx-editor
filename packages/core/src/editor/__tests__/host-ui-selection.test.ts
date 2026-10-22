// HOST UI INSIDE THE PAGES KEEPS ITS OWN SELECTION.
//
// A host can paint interactive widgets inside the page content: a form beside a content
// control, for example. A click there arms a selection gesture, and the browser then reports a
// selection on the widget container. That selection maps to no model position, and the mirror
// used to collapse the document range and write the caret back, which also stole focus from
// the widget's text field. `data-docx-host-ui` marks the subtree as host-owned.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { domSelectionInHostUi, isInHostUi } from '../dom-selection.ts';
import { mountPaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

afterEach(() => {
  document.getSelection()?.removeAllRanges();
});

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

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">alpha</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">beta</w:t></w:r></w:p>';

function mounted() {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, docx(BODY), { scale: 1 });
  if (!opened.ok) throw new Error(opened.reason);
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  const content = pages.querySelector<HTMLElement>('.docx-page-content')!;
  const ui = document.createElement('div');
  ui.setAttribute('data-docx-host-ui', '');
  ui.contentEditable = 'false';
  const field = document.createElement('textarea');
  field.value = 'widget words';
  ui.append(field);
  content.append(ui);
  return { surface: opened.surface, container, pages, ui, field };
}

describe('host UI selections stay outside the document selection', () => {
  test('marker reads text, element, and missing nodes', () => {
    const { surface, container, ui, field } = mounted();
    try {
      const text = document.createTextNode('label');
      ui.append(text);
      expect(isInHostUi(field)).toBe(true);
      expect(isInHostUi(text)).toBe(true);
      expect(isInHostUi(container)).toBe(false);
      expect(isInHostUi(null)).toBe(false);
      expect(domSelectionInHostUi(null)).toBe(false);
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      const range = document.createRange();
      range.setStart(ui, 0);
      range.collapse(true);
      selection.addRange(range);
      expect(domSelectionInHostUi(selection)).toBe(true);
    } finally {
      surface.destroy();
      container.remove();
    }
  });

  test('a gesture in host UI neither collapses the document range nor steals focus', async () => {
    const { surface, container, pages, ui, field } = mounted();
    try {
      const [first, second] = surface.session.paragraphIds();
      const range = {
        anchor: { paragraphId: first!, offset: 1 },
        head: { paragraphId: second!, offset: 3 },
      };
      surface.setSelection(range);
      const revision = surface.session.packageRevision();
      // The browser order: pointerdown arms a gesture, focus moves, selectionchange arrives.
      field.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
      field.focus();
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      const dom = document.createRange();
      dom.setStart(ui, 0);
      dom.collapse(true);
      selection.addRange(dom);
      document.dispatchEvent(new Event('selectionchange'));
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(surface.state().selection).toEqual(range);
      expect(document.activeElement).toBe(field);
      expect(selection.anchorNode).toBe(ui);
      expect(surface.session.packageRevision()).toBe(revision);
      expect(pages.contains(field)).toBe(true);
    } finally {
      surface.destroy();
      container.remove();
    }
  });

  test('the same press and selection without the marker still move the model', async () => {
    const { surface, container, ui, field } = mounted();
    try {
      ui.removeAttribute('data-docx-host-ui');
      const [first, second] = surface.session.paragraphIds();
      const range = {
        anchor: { paragraphId: first!, offset: 1 },
        head: { paragraphId: second!, offset: 3 },
      };
      surface.setSelection(range);
      field.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      const dom = document.createRange();
      dom.setStart(ui, 0);
      dom.collapse(true);
      selection.addRange(dom);
      document.dispatchEvent(new Event('selectionchange'));
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(surface.state().selection).not.toEqual(range);
    } finally {
      surface.destroy();
      container.remove();
    }
  });
});
