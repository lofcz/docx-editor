import { afterEach, expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { blankDocumentBytes } from '@docx-editor.dev/core';
import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import { createMdxComponentRegistry, type MdxComponentRegistry } from './contract';
import { activeMdxRegistry, configureMdxRegistry, withMdxRegistry } from './active';
import { sciobotMdxRegistry } from './sciobot';
import { mdxProps } from './mdx-read';
import { insertBlock } from '../blocks/insert';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { projectBlocks } from '../projection/blocks';
import type { JsonObject } from '../data/json';

function acme(): MdxComponentRegistry {
  const result = createMdxComponentRegistry({
    storage: {
      tagPrefix: 'acme:1:',
      namespace: 'urn:acme:docx:components:1',
      root: 'components',
      nodePrefix: 'acme_',
      reviewNodePrefix: 'acme-review_',
      labelTag: 'acme:label',
      adapterPrefix: 'acme',
      adapterNamespace: 'urn:acme:docx',
    },
    components: [
      {
        type: 'card',
        kind: 'block',
        content: 'native',
        mdx: {
          tag: 'Card',
          kind: 'container',
          attributes: (payload) =>
            Object.entries(payload.attributes).map(([property, value]) => ({
              name: property,
              property,
              value,
            })),
        },
        parse: { tag: 'Card', read: (attributes) => mdxProps(attributes) as JsonObject },
      },
    ],
  });
  if (!result.ok) throw new Error(result.detail);
  return result.registry;
}

afterEach(() => configureMdxRegistry(sciobotMdxRegistry()));

test('the Sciobot registration applies until an application configures its own', () => {
  expect(activeMdxRegistry()).toBe(sciobotMdxRegistry());
  const own = acme();
  configureMdxRegistry(own);
  expect(activeMdxRegistry()).toBe(own);
  expect(withMdxRegistry(sciobotMdxRegistry(), () => activeMdxRegistry())).toBe(
    sciobotMdxRegistry()
  );
  expect(activeMdxRegistry()).toBe(own);
});

test('a configured registry inserts, stores, indexes, and projects its own component', async () => {
  configureMdxRegistry(acme());
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: blankDocumentBytes(), author: 'Ann' });
  try {
    const inserted = insertBlock(editor, {
      parentControlId: null,
      beforeId: null,
      payload: { version: 1, id: 'card-1', type: 'card', attributes: { title: 'Plan' } },
      label: 'Card',
      generatedLabel: 'Plan',
      paragraphs: ['Body'],
    });
    if (!inserted.success) throw new Error(inserted.error);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const pkg = reopened.package;
    const store = readPayloadStore(pkg);
    if (store.status !== 'ready') throw new Error(store.status);
    expect(store.entries.get('card-1')?.label).toBe('Card');
    const block = buildBlockIndex(pkg).byId.get('card-1');
    expect(block?.payload.attributes).toEqual({ title: 'Plan' });
    const session = editor.surface!.session;
    const opened = openDocumentForExport(session, { displayMode: 'all-markup' });
    if (!opened.ok) throw new Error(opened.reason);
    try {
      const projected = projectBlocks(session.currentPackage(), await opened.session.layout(), {
        documentId: 'acme',
        epoch: 'load-1',
        revision: session.packageRevision(),
      });
      if (!projected.success) throw new Error(projected.error);
      expect(projected.projection.text).toContain('<Card title="Plan">');
      expect(projected.projection.text).toContain('Body');
    } finally {
      opened.session.dispose();
    }
    // The Sciobot names find nothing of this embedder's.
    withMdxRegistry(sciobotMdxRegistry(), () => {
      expect(readPayloadStore(pkg).status).toBe('absent');
      expect(buildBlockIndex(pkg).byId.size).toBe(0);
    });
  } finally {
    editor.destroy();
    container.remove();
  }
});
