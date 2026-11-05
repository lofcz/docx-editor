/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import {
  DOCUMENT_COLLABORATION_VERSIONS as versions,
  assertDocumentCollaborationCompatibility as assertCompatible,
} from '../index.ts';
import { CollaborationSchemaError } from '../errors.ts';

test('public full-document contract matches JSON handshake', () => {
  expect(versions.sharedSchemaVersion).toBe(10);
  expect(Object.isFrozen(versions)).toBe(true);
  expect(() => assertCompatible(JSON.parse(JSON.stringify(versions)))).not.toThrow();
  expect(() => assertCompatible({ ...versions, futureCapability: true })).not.toThrow();
});

for (const field of Object.keys(versions) as (keyof typeof versions)[]) {
  for (const received of [0, versions[field] + 1, undefined, '3', null, NaN, Infinity, {}, []]) {
    test(`refuses ${field}=${String(received)} with stable code and bounded diagnostic`, () => {
      const input = Object.freeze({ ...versions, [field]: received });
      let failure: unknown;
      try {
        assertCompatible(input);
      } catch (error) {
        failure = error;
      }
      expect(failure).toBeInstanceOf(CollaborationSchemaError);
      expect(failure).toMatchObject({
        code: field === 'protocolVersion' ? 'protocol-version-mismatch' : 'schema-version-mismatch',
      });
      expect((failure as CollaborationSchemaError).detail).toContain(
        `${field}: expected ${versions[field]}, received `
      );
      expect((failure as CollaborationSchemaError).detail!.length).toBeLessThan(100);
    });
  }
}

test('refuses missing descriptors and inherited claims without coercion or leaking metadata', () => {
  for (const input of [undefined, null, true, 3, '3', [], Object.create(versions)]) {
    expect(() => assertCompatible(input)).toThrow('protocolVersion: expected 1, received missing');
  }
  expect(() =>
    assertCompatible({ ...versions, sharedSchemaVersion: '<script>private data</script>' })
  ).toThrow('sharedSchemaVersion: expected 10, received invalid');
});

test('the reconciled contract refuses earlier fork and upstream rooms', () => {
  expect(versions).toEqual({
    protocolVersion: 1,
    sharedSchemaVersion: 10,
    repairVersion: 17,
    canonicalModelVersion: 1,
  });
  expect(() => assertCompatible({ ...versions, repairVersion: 16 })).toThrow(
    'repairVersion: expected 17, received 16'
  );
  expect(() => assertCompatible({ ...versions, sharedSchemaVersion: 5, repairVersion: 3 })).toThrow(
    'sharedSchemaVersion: expected 10, received 5'
  );
});

test('schema 10 refuses schema 9 rooms until DOCX export and reseeding', () => {
  expect(() => assertCompatible({ ...versions, sharedSchemaVersion: 9 })).toThrow(
    'sharedSchemaVersion: expected 10, received 9'
  );
});
