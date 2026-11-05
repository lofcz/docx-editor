import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  contentControlTextOf,
  findNode,
  isAtomicContentControl,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { mathDocument, importedMath } from '../../test/support/math-document';
import { prepareNativeDocument } from '../native/prepare-document';
import { buildBlockIndex } from './block-index';
import { insertBlock } from './insert';
import { createMathPayload } from './math-codec';

for (const source of ['', 'first\r\n\r\nlast\r\n', '\n\n', 'a\rb\nc'])
  test(`native equation migration keeps exact source and one Word fallback record: ${JSON.stringify(source)}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: mathDocument(),
    });
    try {
      const imported = importedMath(editor.surface!.session.currentPackage());
      expect(
        insertBlock(editor, {
          parentControlId: null,
          beforeId: null,
          payload: createMathPayload('legacy', 'block', source),
          label: 'Equation',
          paragraphs: source ? source.split(/\r\n|\r|\n/) : ['□'],
        }).success
      ).toBe(true);
      const legacy = new Uint8Array(await editor.save());
      const prepared = prepareNativeDocument(legacy);
      if (!prepared.success) throw Error(prepared.error);
      expect(prepared.changed).toBe(true);
      const opened = readOoxmlPackage(prepared.document);
      if (!opened.ok) throw Error(opened.reason);
      const current = buildBlockIndex(opened.package).byId.get('legacy')!;
      const control = findNode(opened.package.parts.get(current.partName)!, current.controlId)!;
      expect(isAtomicContentControl(control)).toBe(true);
      expect(current.payload.attributes.texExpression).toBe(source);
      expect(contentControlTextOf(control)).toBe(
        source ? source.split(/\r\n|\r|\n/).join(' ') : '□'
      );
      const content =
        control.kind === 'contentControl' &&
        control.children.find((child) => child.kind === 'contentControlContent');
      expect(
        content && content.children.filter((child) => child.kind === 'paragraph')
      ).toHaveLength(1);
      expect(importedMath(opened.package)).toEqual(imported);
      expect(prepareNativeDocument(prepared.document)).toMatchObject({
        success: true,
        changed: false,
      });
    } finally {
      editor.destroy();
    }
  });
