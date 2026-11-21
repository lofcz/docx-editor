/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/docx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { contentControlsIn } from '@docx-editor.dev/core/store';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdf } from '../src/index.ts';
import { docx, paragraph } from './fixture.ts';

test('a saved surface with an initial zero-height fold prints the complete Word content', async () => {
  const input = docx(
    paragraph('Heading') +
      '<w:sdt><w:sdtPr/><w:sdtContent>' +
      paragraph('Answer') +
      '</w:sdtContent></w:sdt>'
  );
  const container = document.createElement('div');
  const editor = createDocxEditor({
    container,
    document: input,
    configureInitialBlockView(pkg, view) {
      const [owner] = contentControlsIn(pkg.parts.get(pkg.mainDocumentPart)!.root);
      expect(
        view.setCollapsedContentControls(new Map([[owner!.node.id, { placeholderHeightPt: 0 }]]))
      ).toEqual({ ok: true });
    },
  });
  try {
    expect(container.textContent).not.toContain('Answer');
    const result = await exportPdf(new Uint8Array(await editor.save()), { useSystemFonts: false });
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    try {
      expect(pdf.numPages).toBe(1);
      const text = await (await pdf.getPage(1)).getTextContent();
      expect(
        text.items
          .filter((item) => 'str' in item)
          .map((item) => item.str)
          .join('')
      ).toBe('HeadingAnswer');
    } finally {
      await pdf.loadingTask.destroy();
    }
  } finally {
    editor.destroy();
  }
});
