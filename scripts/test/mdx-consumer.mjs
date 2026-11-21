// Copied into the packed consumer's Vite project. Every import resolves from installed tarballs.
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { contentControlTextOf } from '@docx-editor.dev/core/store';
import {
  createMdxComponentRegistry,
  exportMdxDocument,
  importMdxDocument,
  mountMdxComponents,
  mountMdxDisclosures,
} from '@docx-editor.dev/docx-to-markdown/mdx';

export async function inspectMdxHost({ project, containerTag, action, caption }) {
  const registry = createMdxComponentRegistry({
    storage: {
      tagPrefix: project + ':1:',
      namespace: 'urn:' + project + ':records',
      root: 'records',
      nodePrefix: project + '_',
      reviewNodePrefix: project + '-review_',
      labelTag: project + ':label',
      adapterPrefix: project,
      adapterNamespace: 'urn:' + project + ':word',
    },
    components: [
      {
        type: 'reference',
        kind: 'block',
        content: 'none',
        parse: { tag: 'Reference', read: () => ({}) },
        wordFallback: () => ({ lines: [caption] }),
      },
    ],
  });
  if (!registry.ok) throw new Error(registry.detail);
  const source = [
    'import instrument from "./absent.mjs"',
    '',
    '<details open>',
    '<summary>',
    '',
    'Report',
    '',
    '- Label $x$',
    '',
    '</summary>',
    '',
    '<' + containerTag + '>',
    '',
    '<Reference />',
    '',
    '| Sample | Value |',
    '| --- | --- |',
    '| <Unknown>Depth</Unknown> | $y$ |',
    '',
    '</' + containerTag + '>',
    '',
    '</details>',
  ].join('\n');
  const bytes = importMdxDocument(source, { registry: registry.registry });
  const original = exportMdxDocument(bytes);
  const container = document.createElement('section');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: bytes, zoom: 1 });
  const stopDisclosure = mountMdxDisclosures(editor);
  let observed = '';
  let clicks = 0;
  const stopViews = mountMdxComponents(
    editor,
    new Map([
      [
        containerTag,
        ({ element, observeContent }) => {
          const button = element.ownerDocument.createElement('button');
          button.type = 'button';
          button.textContent = action;
          button.style.pointerEvents = 'auto';
          button.onclick = () => ++clicks;
          element.append(button);
          return observeContent((control) => {
            observed = contentControlTextOf(control);
          });
        },
      ],
    ])
  );
  try {
    const surface = editor.surface;
    surface.flushPendingInput();
    const selector = '[data-mdx-component="' + containerTag + '"] button';
    const button = container.querySelector(selector);
    button.click();
    const match = editor.findMatches('Depth')[0];
    const point = { paragraphId: match.blockId, offset: match.start + 1 };
    surface.setSelection({ anchor: point, head: point });
    if (!editor.exec({ type: 'insertText', text: '!' }).ok) throw new Error('Edit refused');
    surface.flushPendingInput();
    const retained = button === container.querySelector(selector);
    container.querySelector('button[aria-label="Report"]').click();
    surface.flushPendingInput();
    const closed = !container.contains(button);
    container.querySelector('button[aria-label="Report"]').click();
    surface.flushPendingInput();
    const fading = Array.from(
      container.querySelectorAll('.docx-paragraph-fragment, .docx-table-fragment')
    ).filter((element) =>
      element
        .getAnimations()
        .some((animation) =>
          animation.effect?.getKeyframes().some((frame) => frame.opacity !== undefined)
        )
    );
    return {
      clicks,
      action: button.textContent,
      retained,
      closed,
      animatedBody: fading.some((element) => element.textContent.includes('D!epth')),
      animatedSummary: fading.some((element) => element.textContent.includes('Label')),
      childText: observed.includes('D!epth'),
      registeredFallback: observed.includes(caption),
      roundTrip:
        exportMdxDocument(new Uint8Array(await editor.save())) ===
        original.replace('Depth', 'D!epth'),
    };
  } finally {
    stopViews();
    stopDisclosure();
    editor.destroy();
    container.remove();
  }
}
