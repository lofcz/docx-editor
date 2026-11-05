import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { assertNativePropertyProfile } from './property-review-scope';
import { hasEvaluatedNativeProfile } from './native-profile';

const evaluated = {
  protocolVersion: 1,
  sharedSchemaVersion: 10,
  repairVersion: 17,
  canonicalModelVersion: 2,
};

test('property ownership admits only the evaluated native profile without changing room metadata', () => {
  const doc = new Y.Doc();
  try {
    expect(() => assertNativePropertyProfile(doc)).toThrow(
      'Unsupported native property ownership profile.'
    );
    expect(doc.share.size).toBe(0);
    const metadata = doc.getMap('docx-package-meta-v1');
    for (const [key, value] of Object.entries(evaluated)) metadata.set(key, value);
    expect(() => assertNativePropertyProfile(doc)).not.toThrow();
    for (const [key, value] of Object.entries(evaluated)) {
      for (const invalid of [value - 1, value + 1, String(value), null, undefined]) {
        if (invalid === undefined) metadata.delete(key);
        else metadata.set(key, invalid);
        const before = Y.encodeStateAsUpdate(doc);
        expect(() => assertNativePropertyProfile(doc)).toThrow(
          'Unsupported native property ownership profile.'
        );
        expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
        metadata.set(key, value);
      }
    }
  } finally {
    doc.destroy();
  }
});

test('native profile reads refuse wrong root types and resolve decoded map roots without writing', () => {
  const malformed = new Y.Doc();
  const source = new Y.Doc();
  const decoded = new Y.Doc();
  try {
    malformed.getArray('docx-package-meta-v1').push([evaluated]);
    const malformedBefore = Y.encodeStateAsUpdate(malformed);
    expect(hasEvaluatedNativeProfile(malformed)).toBe(false);
    expect(() => assertNativePropertyProfile(malformed)).toThrow(
      'Unsupported native property ownership profile.'
    );
    expect(Y.encodeStateAsUpdate(malformed)).toEqual(malformedBefore);
    const metadata = source.getMap('docx-package-meta-v1');
    for (const [key, value] of Object.entries(evaluated)) metadata.set(key, value);
    Y.applyUpdate(decoded, Y.encodeStateAsUpdate(source));
    const decodedBefore = Y.encodeStateAsUpdate(decoded);
    expect(hasEvaluatedNativeProfile(decoded)).toBe(true);
    expect(Y.encodeStateAsUpdate(decoded)).toEqual(decodedBefore);
  } finally {
    malformed.destroy();
    source.destroy();
    decoded.destroy();
  }
});
