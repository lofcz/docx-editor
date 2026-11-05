import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { assertNativePropertyProfile } from './property-review-scope';

const evaluated = {
  protocolVersion: 1,
  sharedSchemaVersion: 10,
  repairVersion: 17,
  canonicalModelVersion: 1,
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
