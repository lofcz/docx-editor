import { expect, test } from 'bun:test';
import { createMdxComponentRegistry, type MdxComponentRegistry } from './contract';
import { mdxProps } from './mdx-read';
import { sciobotMdxRegistry } from './sciobot';
import { domainNodeFromAst } from '../diff/domain-ast';
import { insertedPayload } from '../diff/inserted-blocks';
import { payloadMatches } from '../diff/inserted-cell-fragments';
import { domainDiffStrategy, sameContainerIdentity } from '../diff/domain-strategies';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import type { JsonObject } from '../data/json';

function registry(): MdxComponentRegistry {
  const result = createMdxComponentRegistry({
    storage: {
      tagPrefix: 'acme:1:',
      namespace: 'urn:acme:docx:components:1',
      root: 'components',
      nodePrefix: 'acme_',
      reviewNodePrefix: 'acme-review_',
      labelTag: 'acme:label',
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
              name: property === 'tone' ? 'variant' : property,
              property,
              value,
            })),
        },
        parse: {
          tag: 'Card',
          read: (attributes) => mdxProps(attributes, { rename: { variant: 'tone' } }) as JsonObject,
        },
        diff: { kind: 'container', identityProps: ['title'] },
      },
    ],
  });
  if (!result.ok) throw new Error(result.detail);
  return result.registry;
}

function firstElement(markdown: string): TargetAstNode {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.ast.children[0]!;
}

test('an embedder component parses, binds, and matches through its own registry', () => {
  const acme = registry();
  const ast = firstElement('<Card title="Plan" variant="calm">\n\nBody\n\n</Card>');
  expect(domainNodeFromAst(ast, acme)).toEqual({ title: 'Plan', tone: 'calm', type: 'card' });
  const payload = insertedPayload(ast, 'card-1', acme);
  expect(payload).toEqual({
    version: 1,
    id: 'card-1',
    type: 'card',
    attributes: { title: 'Plan', tone: 'calm' },
  });
  expect(payloadMatches(ast, payload!, acme)).toBe(true);
  const changed = firstElement('<Card title="Plan" variant="loud">\n\nBody\n\n</Card>');
  expect(payloadMatches(changed, payload!, acme)).toBe(false);
  expect(domainDiffStrategy({ type: 'card' }, acme)).toEqual({
    kind: 'container',
    identityProps: ['title'],
  });
  expect(
    sameContainerIdentity(
      { type: 'card', title: 'Plan', tone: 'calm' },
      { type: 'card', title: 'Plan', tone: 'loud' },
      acme
    )
  ).toBe(true);
});

test('each registry reads only its own components', () => {
  const acme = registry();
  const callout = firstElement('<callout type="info">\n\nText\n\n</callout>');
  const card = firstElement('<Card title="Plan">\n\nBody\n\n</Card>');
  expect(domainNodeFromAst(callout, acme)).toBeNull();
  expect(insertedPayload(callout, 'x', acme)).toBeNull();
  expect(domainNodeFromAst(card, sciobotMdxRegistry())).toBeNull();
  expect(domainNodeFromAst(callout, sciobotMdxRegistry())).toEqual({
    variant: 'info',
    type: 'callout',
  });
  // Fenced code is a Sciobot component; another embedder that registers none leaves it plain.
  const code = firstElement('```js\nx\n```');
  expect(insertedPayload(code, 'c', acme)).toBeNull();
  expect(insertedPayload(code, 'c', sciobotMdxRegistry())?.type).toBe('code_block');
});
