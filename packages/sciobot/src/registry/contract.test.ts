import { expect, test } from 'bun:test';
import {
  createMdxComponentRegistry,
  type MdxComponentDefinition,
  type MdxComponentStorage,
  type MdxOpaqueTags,
} from './contract';
import { sciobotMdxRegistry } from './sciobot';
import { DOMAIN_BLOCK_TAGS, GENERATED_LABEL_TAG } from '../projection/block-tags';
import { SCIOBOT_DIFF_STRATEGIES } from './sciobot-diff';
import { SCIOBOT_STORAGE } from './storage';
import { SCIOBOT_PAYLOAD_STORAGE } from '../blocks/payload-opc-codec';

const storage: MdxComponentStorage = {
  tagPrefix: 'acme:1:',
  namespace: 'urn:acme:docx:components:1',
  root: 'components',
  nodePrefix: 'acme_',
  reviewNodePrefix: 'acme-review_',
  labelTag: 'acme:label',
  adapterPrefix: 'acme',
  adapterNamespace: 'urn:acme:docx',
};
const card: MdxComponentDefinition = {
  type: 'card',
  kind: 'block',
  content: 'native',
  mdx: { tag: 'Card', kind: 'container', attributes: () => [] },
  diff: { kind: 'container', identityProps: ['title'] },
};
const badge: MdxComponentDefinition = {
  type: 'badge',
  kind: 'inline',
  content: 'none',
  mdx: { tag: 'Badge', kind: 'atomic', attributes: () => [] },
};

test('a registry indexes components, MDX mappings, and diff strategies by payload type', () => {
  const result = createMdxComponentRegistry({
    storage,
    components: [card, badge],
    opaque: { allowed: { Chart: ['data'] }, literal: ['Chart'], dataOnly: [] },
  });
  if (!result.ok) throw new Error(result.detail);
  const registry = result.registry;
  expect([...registry.components.keys()]).toEqual(['card', 'badge']);
  expect(registry.blockTags.get('card')?.tag).toBe('Card');
  expect(registry.diffStrategies.get('card')).toEqual(card.diff);
  expect(registry.diffStrategies.has('badge')).toBe(false);
  expect(registry.opaque.allowed.get('Chart')).toEqual(['data']);
  expect(registry.opaque.literal.has('Chart')).toBe(true);
  expect(Object.isFrozen(registry)).toBe(true);
  expect(Object.isFrozen(registry.storage)).toBe(true);
});

test('invalid storage names refuse before any component is read', () => {
  for (const [field, value] of [
    ['tagPrefix', 'acme'],
    ['tagPrefix', '__proto__'],
    ['namespace', 'javascript:alert(1)'],
    ['namespace', 'urn:acme: spaced'],
    ['root', 'constructor'],
    ['nodePrefix', 'acme'],
    ['reviewNodePrefix', 'acme_'],
    ['labelTag', 'acme:1:label'],
    ['adapterPrefix', 'Acme:'],
    ['adapterNamespace', 'urn:acme:docx:components:1'],
  ] as const) {
    const result = createMdxComponentRegistry({
      storage: { ...storage, [field]: value },
      components: [card],
    });
    expect(result).toMatchObject({ ok: false, reason: 'invalid-storage', detail: field });
  }
});

test('hostile, duplicate, or inconsistent components refuse with a stable reason', () => {
  const cases: [MdxComponentDefinition[], string][] = [
    [[{ ...card, type: '__proto__' }], 'invalid-component'],
    [[{ ...card, type: 'Card' }], 'invalid-component'],
    [[{ ...card, type: 'x'.repeat(65) }], 'invalid-component'],
    [[{ ...card, mdx: { ...card.mdx!, tag: 'bad tag' } }], 'invalid-component'],
    [[{ ...card, content: 'none' }], 'invalid-component'],
    [[{ ...badge, content: 'native' }], 'invalid-component'],
    [
      [{ ...card, diff: { kind: 'container', identityProps: ['constructor'] } }],
      'invalid-component',
    ],
    [[card, card], 'duplicate-component'],
    [[card, { ...card, type: 'other_card' }], 'duplicate-tag'],
  ];
  for (const [components, reason] of cases)
    expect(createMdxComponentRegistry({ storage, components })).toMatchObject({
      ok: false,
      reason,
    });
  // One tag may name an atom and a container, as a citation and a citation excerpt do.
  expect(
    createMdxComponentRegistry({
      storage,
      components: [card, { ...badge, mdx: { ...badge.mdx!, tag: 'Card' } }],
    }).ok
  ).toBe(true);
});

test('opaque tags must be allowed before they can be literal or data-only', () => {
  for (const opaque of <MdxOpaqueTags[]>[
    { allowed: { 'bad tag': [] }, literal: [], dataOnly: [] },
    { allowed: { Chart: ['__proto__'] }, literal: [], dataOnly: [] },
    { allowed: {}, literal: ['Chart'], dataOnly: [] },
    { allowed: {}, literal: [], dataOnly: ['Chart'] },
  ])
    expect(createMdxComponentRegistry({ storage, components: [], opaque })).toMatchObject({
      ok: false,
      reason: 'invalid-opaque-tag',
    });
});

test('the Sciobot registry keeps the stored names and every existing mapping', () => {
  const registry = sciobotMdxRegistry();
  expect(registry.storage).toEqual({
    tagPrefix: 'scio:1:',
    namespace: 'urn:sciobot:docx:blocks:1',
    root: 'blocks',
    nodePrefix: 'scio_',
    reviewNodePrefix: 'scio-review_',
    labelTag: 'scio:label',
    adapterPrefix: 'scio',
    adapterNamespace: 'urn:sciobot:docx',
  });
  expect([...registry.blockTags]).toEqual([...DOMAIN_BLOCK_TAGS]);
  for (const type of [...registry.components.keys()])
    expect(registry.diffStrategies.get(type)).toEqual(SCIOBOT_DIFF_STRATEGIES[type]);
  expect(registry.components.get('listening')).toMatchObject({ content: 'opaque-mdx' });
  expect(registry.components.get('mention')).toMatchObject({ kind: 'inline', content: 'none' });
  expect(registry.components.get('callout')).toMatchObject({ kind: 'block', content: 'native' });
  expect(registry.opaque.dataOnly.has('listening')).toBe(true);
});

test('the dependency-free Sciobot storage names equal the codec constants', () => {
  expect(SCIOBOT_STORAGE).toEqual({
    ...SCIOBOT_PAYLOAD_STORAGE,
    labelTag: GENERATED_LABEL_TAG,
    adapterPrefix: 'scio',
    adapterNamespace: 'urn:sciobot:docx',
  });
  expect(sciobotMdxRegistry().storage).toEqual(SCIOBOT_STORAGE);
});
