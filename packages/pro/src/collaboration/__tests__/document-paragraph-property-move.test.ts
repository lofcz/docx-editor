/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import {
  moveNode,
  withPart,
  writeOoxmlPackage,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { zipDocument } from './document-peer-support';
import {
  seedReplica,
  joinReplica,
  destroyReplica,
  packageOf,
  packageFingerprint,
  collectKind,
  loadPackage,
  applyJournal,
} from './document-support';
import { idOf } from '../document/identity';

for (const destinationIndex of [0, 1])
  test(`paragraph property moves stay in records at native child index ${destinationIndex}`, async () => {
    const replica = await seedReplica(
      loadPackage(
        zipDocument(
          '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Beta</w:t></w:r></w:p>'
        )
      )
    );
    try {
      const before = packageOf(replica),
        [source, target] = collectKind(before, 'paragraph');
      const props = source!.children.find((n) => n.kind === 'paragraphProperties')!;
      const part = before.parts.get(before.mainDocumentPart)!;
      const direct = moveNode(part, props.id, target!.id, 0);
      if (!direct.ok) throw Error('Independent native move refused.');
      const expected = withPart(before, direct.part);
      const journal = {
        effects: [
          {
            kind: 'moveNode' as const,
            logicalId: idOf(props),
            destinationParentLogicalId: idOf(target!),
            destinationIndex,
          },
        ],
      };
      applyJournal(replica, journal);
      expect(packageFingerprint(packageOf(replica))).toBe(packageFingerprint(expected));
      applyJournal(replica, journal);
      expect(packageFingerprint(packageOf(replica))).toBe(packageFingerprint(expected));
      const reopened = readOoxmlPackage(writeOoxmlPackage(packageOf(replica)));
      if (!reopened.ok) throw Error(reopened.reason);
      expect(packageFingerprint(reopened.package)).toBe(packageFingerprint(expected));
      const cold = joinReplica(replica);
      try {
        expect(packageFingerprint(packageOf(cold))).toBe(packageFingerprint(expected));
      } finally {
        destroyReplica(cold);
      }
    } finally {
      destroyReplica(replica);
    }
  });
