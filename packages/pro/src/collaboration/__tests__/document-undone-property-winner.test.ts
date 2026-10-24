/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { replaceNode, withPart } from '@docx-editor.dev/core/store';
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
  WML,
} from './document-support';
import { idOf } from '../document/identity';
import { NODE_UNDONE_FIELD } from '../document/schema';

test('an empty undone constructor cannot hide a surviving paragraph mark property holder', async () => {
  const replica = await seedReplica(
    loadPackage(
      zipDocument(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:pPr><w:rPr><w:ins w:id="7" w:author="AI"/></w:rPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>'
      )
    )
  );
  try {
    const before = packageOf(replica),
      fingerprint = packageFingerprint(before);
    const props = collectKind(before, 'paragraphProperties')[0]!,
      constructor = replica.mint.take();
    applyJournal(replica, {
      effects: [
        {
          kind: 'putNode',
          descriptor: {
            logicalId: constructor,
            kind: 'runProperties',
            qname: { namespaceUri: WML, localName: 'rPr', prefix: 'w' },
          },
        },
        {
          kind: 'spliceChildren',
          parentLogicalId: idOf(props),
          start: 0,
          deleteCount: 0,
          childLogicalIds: [constructor],
        },
      ],
    });
    replica.registry.schema.nodes.get(constructor)!.set(NODE_UNDONE_FIELD, true);
    expect(replica.materializer.rebuild().ok).toBe(true);
    expect(packageFingerprint(packageOf(replica))).toBe(fingerprint);
    const marker = collectKind(before, 'generic').find((node) => node.localName === 'ins')!;
    if (marker.kind !== 'generic') throw Error('Missing native paragraph mark.');
    const sourcePart = before.parts.get(before.mainDocumentPart)!;
    const changed = replaceNode(sourcePart, marker.id, {
      ...marker,
      attributes: marker.attributes.map((attr) =>
        attr.namespaceUri === WML && attr.localName === 'author'
          ? { ...attr, value: 'Reviewer' }
          : attr
      ),
    });
    if (!changed.ok) throw Error('Independent mark edit refused.');
    applyJournal(replica, {
      effects: [
        {
          kind: 'setAttribute',
          logicalId: idOf(marker),
          qname: { namespaceUri: WML, localName: 'author', prefix: 'w' },
          value: 'Reviewer',
        },
      ],
    });
    const edited = packageFingerprint(withPart(before, changed.part));
    expect(packageFingerprint(packageOf(replica))).toBe(edited);
    for (const undone of [false, true]) {
      replica.registry.schema.nodes.get(constructor)!.set(NODE_UNDONE_FIELD, undone);
      expect(replica.materializer.rebuild().ok).toBe(true);
      expect(packageFingerprint(packageOf(replica))).toBe(edited);
    }
    const cold = joinReplica(replica);
    try {
      expect(packageFingerprint(packageOf(cold))).toBe(edited);
    } finally {
      destroyReplica(cold);
    }
  } finally {
    destroyReplica(replica);
  }
});
