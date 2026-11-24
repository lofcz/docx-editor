/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as Y from 'yjs';
import {
  assertCollaborationFormatCompatibility,
  COLLABORATION_FORMAT_VERSION,
  readCollaborationFormatVersion,
} from '../document-format.ts';
import { readCollaborationDocument } from '../document-read.ts';
import { CollaborationSchemaError } from '../errors.ts';
import { collaborationMigrationNeed, migrateCollaborationRoom } from '../room-migration.ts';
import { loadPackage, packageFingerprint } from './document-support.ts';

// A repair-19 provider and two peers created this room and exported the indented paragraph.
const priorState = readFileSync(new URL('./fixtures/repair19-history.state', import.meta.url));
const priorExport = readFileSync(new URL('./fixtures/repair19-history.docx', import.meta.url));
const priorFormat = 'docx-collaboration:1.10.19.2';

test('repair 20 refuses repair-19 peers and preserves the original room on refusal', () => {
  expect(COLLABORATION_FORMAT_VERSION).toBe('docx-collaboration:1.10.20.2');
  expect(() => assertCollaborationFormatCompatibility(priorFormat)).toThrow(
    CollaborationSchemaError
  );
  const room = new Y.Doc();
  try {
    Y.applyUpdate(room, priorState);
    expect(readCollaborationFormatVersion(room)).toBe(priorFormat);
    const before = Y.encodeStateAsUpdate(room);
    expect(() => readCollaborationDocument(room)).toThrow(CollaborationSchemaError);
    expect(Y.encodeStateAsUpdate(room)).toEqual(before);
  } finally {
    room.destroy();
  }
});

test('the verified repair-19 export seeds a fresh repair-20 room with the same document', async () => {
  expect(collaborationMigrationNeed(priorState)).toBe('earlier-format');
  const migrated = await migrateCollaborationRoom({ state: priorState, exported: priorExport });
  if (!migrated.ok) throw Error(migrated.reason);
  const room = new Y.Doc();
  try {
    Y.applyUpdate(room, migrated.state);
    expect(collaborationMigrationNeed(migrated.state)).toBe('current');
    expect(readCollaborationFormatVersion(room)).toBe(COLLABORATION_FORMAT_VERSION);
    expect(packageFingerprint(loadPackage(readCollaborationDocument(room)))).toEqual(
      packageFingerprint(loadPackage(priorExport))
    );
  } finally {
    room.destroy();
  }
});
