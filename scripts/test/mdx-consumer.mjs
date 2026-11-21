// Copied into the packed consumer's Vite project. Every import resolves from installed tarballs.
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { contentControlTextOf } from '@docx-editor.dev/core/store';
import { exportMdxDocument, importMdxDocument } from '@docx-editor.dev/docx-to-markdown/mdx';
import {
  createMdxComponentRegistry,
  mountMdxComponents,
  mountMdxDisclosures,
} from '@docx-editor.dev/docx-to-markdown/mdx/view';

async function untilFrame(check) {
  const deadline = performance.now() + 5000;
  while (!check()) {
    if (performance.now() > deadline) throw new Error('Disclosure did not settle');
    await new Promise((resolve) => requestAnimationFrame(resolve));
  }
}

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
    '| Report | Unit |',
    '| --- | --- |',
    '| **Maximum** | kPa |',
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
    '',
    '<details open={false}>',
    '<summary>Explicitly closed</summary>',
    '',
    'Closed body',
    '',
    '</details>',
    '',
    '<details>',
    '<summary>Default state</summary>',
    '',
    'Default body',
    '',
    '</details>',
  ].join('\n');
  const bytes = importMdxDocument(source, { registry: registry.registry });
  const expected = exportMdxDocument(
    importMdxDocument(source.replace('Depth', 'D!epth'), { registry: registry.registry })
  );
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
          button.onclick = () => {
            clicks++;
            const match = editor.findMatches('Depth')[0];
            const point = { paragraphId: match.blockId, offset: match.start + 1 };
            editor.surface.setSelection({ anchor: point, head: point });
            if (!editor.exec({ type: 'insertText', text: '!' }).ok) throw new Error('Edit refused');
          };
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
    for (const [label, expanded] of [
      ['Report', 'true'],
      ['Explicitly closed', 'false'],
      ['Default state', 'false'],
    ]) {
      const toggle = container.querySelector('button[aria-label="' + label + '"]');
      if (toggle?.getAttribute('aria-expanded') !== expanded)
        throw new Error('Unexpected initial disclosure state for ' + label);
    }
    await window.verifyTableSummaryKeys();
    const root = document.documentElement;
    const settled = () => !root.matches('.docx-mdx-opening, .docx-mdx-closing');
    await untilFrame(settled);
    const selector = '[data-mdx-component="' + containerTag + '"] button';
    const button = container.querySelector(selector);
    button.click();
    surface.flushPendingInput();
    const retained = button === container.querySelector(selector);
    const toggle = container.querySelector('button[aria-label="Report"]');
    toggle.click();
    await untilFrame(() => !container.contains(button) && settled());
    const closed = !container.querySelector(selector);
    const summaryTop = () => toggle.getBoundingClientRect().top;
    const closedTop = summaryTop();
    let animatedSummary = false;
    let clippedBody = false;
    let translatedBody = false;
    toggle.click();
    await untilFrame(() => {
      animatedSummary ||= Math.abs(summaryTop() - closedTop) > 1;
      const movement = document
        .getAnimations()
        .find((a) => a.animationName === 'docx-mdx-move-following');
      if (movement) {
        clippedBody ||= getComputedStyle(
          root,
          '::view-transition-old(docx-mdx-paper)'
        ).clipPath.startsWith('inset(');
        const frames = movement.effect.getKeyframes();
        translatedBody ||=
          frames.length >= 2 &&
          frames[0].transform !== frames.at(-1).transform &&
          parseFloat(root.style.getPropertyValue('--docx-mdx-travel')) > 0;
      }
      return container.querySelector(selector) && settled();
    });
    return {
      clicks,
      action: button.textContent,
      retained,
      closed,
      clippedBody,
      translatedBody,
      animatedSummary,
      childText: observed.includes('D!epth'),
      registeredFallback: observed.includes(caption),
      roundTrip: exportMdxDocument(new Uint8Array(await editor.save())) === expected,
    };
  } finally {
    stopViews();
    stopDisclosure();
    editor.destroy();
    container.remove();
  }
}
