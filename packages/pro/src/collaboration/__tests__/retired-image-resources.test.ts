/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { canonicalOoxmlFingerprint, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { MemoryBlobStore, type BlobBytesStore } from '../document/seed.ts';
import { asLogicalId } from '../document/identity.ts';
import { relationshipDecodeCount } from '../document/relationship-store.ts';
import { isNodeMap } from '../document/schema.ts';
import { REL, R, zipDocument } from './document-peer-support.ts';
import {
  applyJournal,
  spliceTextJournal,
  destroyReplica,
  loadPackage,
  packageOf,
  packageFingerprint,
  seedReplica,
  joinReplica,
  findText,
  type Replica,
} from './document-support.ts';

const NAME = '/word/media/image.png';
const RELS = '/word/_rels/document.xml.rels';
const OWNER = '/word/document.xml';
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const EXT = 'urn:test:resource-extension';
function fixture(override: boolean) {
  const entries = unzipSync(
    zipDocument(
      `<w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="picture"/></w:drawing></w:r></w:p>`,
      {
        documentRels: `<Relationships xmlns="${REL}" xmlns:x="${EXT}" x:root="keep"><Relationship Id="picture" Type="${R}/image" Target="media/image.png" x:row="keep"/></Relationships>`,
        overrides: override ? `<Override PartName="${NAME}" ContentType="image/png"/>` : '',
      }
    )
  );
  entries['word/media/image.png'] = PNG;
  if (!override)
    entries['[Content_Types].xml'] = strToU8(
      strFromU8(entries['[Content_Types].xml']!).replace(
        '</Types>',
        '<Default Extension="png" ContentType="image/png"/></Types>'
      )
    );
  return loadPackage(zipSync(entries));
}
function retire(replica: Replica, removePart = true) {
  replica.doc.transact(() => {
    for (const descriptor of replica.registry.binaries())
      if (descriptor.storageKey.includes('/media/'))
        replica.registry.deleteBinary(descriptor.storageKey);
    replica.registry.deleteRelationship(OWNER, 'picture');
    replica.registry.deleteContentTypeOverride(NAME);
    if (removePart) replica.registry.deleteXmlPart(RELS);
  });
}
function media(replica: Replica) {
  return [...packageOf(replica).partBytes].filter(([key]) => key.includes('/media/'));
}
for (const override of [false, true]) {
  test(`live references recover exact relationship extensions and MIME declaration mode ${override}`, async () => {
    const original = fixture(override),
      replica = await seedReplica(original);
    try {
      retire(replica);
      const recovered = packageOf(replica);
      expect([...recovered.partBytes].filter(([key]) => key.includes('/media/'))).toEqual([
        [NAME, PNG],
      ]);
      expect(recovered.relationships.get(OWNER)).toEqual(original.relationships.get(OWNER));
      expect(recovered.contentTypes).toEqual(original.contentTypes);
      expect(canonicalOoxmlFingerprint(recovered.parts.get(RELS)!)).toBe(
        canonicalOoxmlFingerprint(original.parts.get(RELS)!)
      );
      expect(packageFingerprint(loadPackage(writeOoxmlPackage(recovered)))).toBe(
        packageFingerprint(loadPackage(writeOoxmlPackage(original)))
      );
      const full = replica.materializer.rebuildFull();
      expect(full.ok).toBe(true);
      if (!full.ok) throw new Error(full.code);
      expect(packageFingerprint(full.package)).toBe(packageFingerprint(recovered));
      expect(full.package.partBytes).toEqual(recovered.partBytes);
      const cold = joinReplica(replica);
      try {
        expect(packageFingerprint(packageOf(cold))).toBe(packageFingerprint(recovered));
        expect(media(cold)).toEqual([[NAME, PNG]]);
      } finally {
        destroyReplica(cold);
      }
    } finally {
      destroyReplica(replica);
    }
  });
}

test('removing the drawing owner releases every recovered resource', async () => {
  const replica = await seedReplica(fixture(true));
  try {
    retire(replica);
    expect(media(replica)).toEqual([[NAME, PNG]]);
    const paragraph = packageOf(replica).parts.get(OWNER)!.root.children[0]!;
    if (paragraph.kind === 'textValue') throw new Error('body missing');
    replica.doc.transact(() =>
      replica.registry.spliceChildren(asLogicalId(paragraph.id), 0, paragraph.children.length, [])
    );
    expect(media(replica)).toEqual([]);
    expect(packageOf(replica).relationships.get(OWNER) ?? []).toEqual([]);
    expect(packageOf(replica).contentTypes.overrides.has(NAME)).toBe(false);
    expect(packageOf(replica).parts.has(RELS)).toBe(false);
  } finally {
    destroyReplica(replica);
  }
});

test('text-only edits reuse recovered bytes without decoding relationships', async () => {
  let reads = 0;
  const inner = new MemoryBlobStore();
  const blobs: BlobBytesStore = {
    get(digest) {
      reads++;
      return inner.get(digest);
    },
    put(digest, bytes) {
      inner.put(digest, bytes);
    },
  };
  const replica = await seedReplica(fixture(true), blobs);
  try {
    retire(replica);
    expect(media(replica)).toEqual([[NAME, PNG]]);
    const before = reads,
      decoded = relationshipDecodeCount();
    for (let i = 0; i < 4; i++) {
      const text = findText(packageOf(replica), 'Hello' + '!'.repeat(i));
      applyJournal(replica, spliceTextJournal(text.id, text.value.length, '!'));
      expect(media(replica)).toEqual([[NAME, PNG]]);
    }
    expect(reads).toBe(before);
    expect(relationshipDecodeCount()).toBe(decoded);
  } finally {
    destroyReplica(replica);
  }
});

test('retired relationship root edits invalidate recovered XML', async () => {
  const replica = await seedReplica(fixture(true));
  try {
    const root = packageOf(replica).parts.get(RELS)!.root;
    retire(replica);
    expect(media(replica)).toEqual([[NAME, PNG]]);
    replica.doc.transact(() =>
      replica.registry.setAttribute(
        asLogicalId(root.id),
        { namespaceUri: EXT, localName: 'root', prefix: 'x' },
        'changed'
      )
    );
    expect(
      packageOf(replica)
        .parts.get(RELS)!
        .root.attributes.find((a) => a.localName === 'root')?.value
    ).toBe('changed');
    expect(media(replica)).toEqual([[NAME, PNG]]);
  } finally {
    destroyReplica(replica);
  }
});

test('missing retained blobs fail materialization', async () => {
  let missing = false;
  const inner = new MemoryBlobStore();
  const blobs: BlobBytesStore = {
    get(digest) {
      return missing ? null : inner.get(digest);
    },
    put(digest, bytes) {
      inner.put(digest, bytes);
    },
  };
  const replica = await seedReplica(fixture(true), blobs);
  try {
    retire(replica);
    missing = true;
    expect(replica.materializer.current()).toMatchObject({ ok: false, code: 'missing-blob' });
  } finally {
    destroyReplica(replica);
  }
});

test.each(['retainedDigest', 'size', 'storageKey'] as const)(
  'invalid retained %s cannot restore binary bytes',
  async (field) => {
    const replica = await seedReplica(fixture(true));
    try {
      retire(replica);
      replica.doc.transact(() => {
        for (const entry of replica.registry.schema.binaries.values())
          if (isNodeMap(entry) && entry.get('digest') === '')
            entry.set(field, field === 'size' ? Number.MAX_SAFE_INTEGER : '../invalid');
      });
      expect(media(replica)).toEqual([]);
      expect(packageOf(replica).relationships.get(OWNER) ?? []).toEqual([]);
    } finally {
      destroyReplica(replica);
    }
  }
);
