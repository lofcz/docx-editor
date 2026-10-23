import { expect, test } from 'bun:test';
import { secureRandomUuid } from '../package/random-uuid.ts';

function withCrypto(value: unknown, run: () => void): void {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'crypto');
  Object.defineProperty(globalThis, 'crypto', { configurable: true, value });
  try {
    run();
  } finally {
    if (original) Object.defineProperty(globalThis, 'crypto', original);
    else Reflect.deleteProperty(globalThis, 'crypto');
  }
}

test('native identities use the host UUID method with its required receiver', () => {
  const source = {
    randomUUID() {
      expect(this).toBe(source);
      return '00000000-0000-4000-8000-000000000001';
    },
  };
  withCrypto(source, () => expect(secureRandomUuid()).toBe('00000000-0000-4000-8000-000000000001'));
});

test('native identities use secure random bytes when randomUUID is unavailable', () => {
  withCrypto({ getRandomValues: (bytes: Uint8Array) => bytes.fill(0xff) }, () => {
    expect(secureRandomUuid()).toBe('ffffffff-ffff-4fff-bfff-ffffffffffff');
  });
});

test('native identities refuse a host without a secure random source', () => {
  withCrypto(undefined, () => expect(secureRandomUuid).toThrow('secure random source'));
});
