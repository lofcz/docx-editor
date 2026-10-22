import { expect, test } from 'bun:test';
import fixture from '../__tests__/fixtures/block-payloads.json';
import { decodeBlockPayload, encodeBlockPayload, type BlockPayload } from './payload-schema';

test('nested blocks, complete ASTs, and unknown fields roundtrip without defaults', () => {
  const raw = JSON.stringify(fixture.nested);
  const result = decodeBlockPayload(raw);
  expect(result.status).toBe('supported');
  if (result.status !== 'supported') throw new Error('Expected supported payload');
  expect<unknown>(result.value).toEqual(fixture.nested);
  expect(JSON.parse(encodeBlockPayload(result.value))).toEqual(fixture.nested);
  expect(Object.hasOwn(result.value.attributes, 'resource')).toBe(false);
});

test('unknown versions remain opaque with their original JSON retained', () => {
  const raw = JSON.stringify(fixture.future, null, 2);
  const result = decodeBlockPayload(raw);
  expect(result.status).toBe('unsupported-version');
  if (result.status !== 'unsupported-version') throw new Error('Expected opaque payload');
  expect(result.raw).toBe(raw);
  expect(result.version).toBe(9);
  expect(result.value).toEqual(fixture.future);
  expect(decodeBlockPayload('{"version":2,"newShape":true}').status).toBe('unsupported-version');
});

test('missing and malformed fields do not become empty valid payloads', () => {
  for (const raw of ['{', 'null', '[]', '{}', '{"id":"a","type":"x","attributes":{}}']) {
    expect(decodeBlockPayload(raw).status).toBe('invalid');
  }
  for (const field of ['id', 'type', 'version', 'attributes']) {
    const value: Record<string, unknown> = { id: 'a', type: 'x', version: 1, attributes: {} };
    delete value[field];
    expect(decodeBlockPayload(JSON.stringify(value)).status).toBe('invalid');
  }
});

test('payload size, node, and depth limits fail before a valid record escapes', () => {
  const raw = JSON.stringify(fixture.nested);
  expect(decodeBlockPayload(raw, { maxBytes: 20 }).status).toBe('invalid');
  expect(decodeBlockPayload(raw, { maxNodes: 3 }).status).toBe('invalid');
  expect(decodeBlockPayload(raw, { maxDepth: 2 }).status).toBe('invalid');
});

test('unknown object keys stay data and never change object prototypes', () => {
  const raw =
    '{"id":"a","type":"constructor","version":1,"attributes":{"__proto__":{"polluted":true},"constructor":"keep"}}';
  const result = decodeBlockPayload(raw);
  expect(result.status).toBe('supported');
  if (result.status !== 'supported') throw new Error('Expected supported data');
  expect(Object.hasOwn(result.value.attributes, '__proto__')).toBe(true);
  expect<unknown>(result.value.attributes.constructor).toBe('keep');
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  expect(JSON.parse(encodeBlockPayload(result.value))).toEqual(JSON.parse(raw));
});

test('encoding rejects values JSON would silently drop or replace', () => {
  const base = { id: 'a', type: 'x', version: 1, attributes: {} } as const;
  for (const value of [undefined, Infinity, () => 'code', new Date()]) {
    const payload = { ...base, extra: value } as unknown as BlockPayload;
    expect(() => encodeBlockPayload(payload)).toThrow();
  }
  let executed = false;
  const payload = {
    ...base,
    toJSON() {
      executed = true;
      return {};
    },
  } as unknown as BlockPayload;
  expect(() => encodeBlockPayload(payload)).toThrow();
  expect(executed).toBe(false);
});

test('encoding rejects cycles and sparse arrays', () => {
  const cycle: Record<string, unknown> = { id: 'a', type: 'x', version: 1, attributes: {} };
  cycle.self = cycle;
  expect(() => encodeBlockPayload(cycle as unknown as BlockPayload)).toThrow('cyclic-payload');
  const sparse = { id: 'a', type: 'x', version: 1, attributes: {}, values: Array(3) };
  expect(() => encodeBlockPayload(sparse as BlockPayload)).toThrow('sparse-or-extended-array');
  const extended = Array(1);
  Object.assign(extended, { arbitrary: 'ignored by JSON' });
  expect(() =>
    encodeBlockPayload({ id: 'a', type: 'x', version: 1, attributes: {}, values: extended })
  ).toThrow('sparse-or-extended-array');
});
