/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  readOoxmlPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  observeCanonicalPrimitiveJournal,
  runObservedStoreTransaction,
} from '../../../../core/src/store/package/canonical-primitive-capture.ts';
import { withNumberingTypeBytes } from '../document/materialize-numbering-types.ts';
import { DocumentRegistry } from '../document/registry.ts';
import {
  isElementRecord,
  NODE_UNDONE_FIELD,
  nodeRecordUndone,
  packageSchemaOf,
} from '../document/schema.ts';
import {
  destroyReplica,
  findText,
  joinReplica,
  packageOf,
  seedReplica,
  shownParentOf,
} from './document-support.ts';
import { zipDocument } from './document-peer-support.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NUMBERING = '/word/numbering.xml';
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const body = '<w:p><w:r><w:t>Body</w:t></w:r></w:p>';
const override = `<Override PartName="${NUMBERING}" ContentType="${TYPE}"/>`;

test('retained numbering overrides reach saved bytes without losing opaque content types or writing a journal', () => {
  const read = readOoxmlPackage(
    zipDocument(body, { overrides: '<x:Ledger xmlns:x="urn:ledger" x:stamp="keep"/>' })
  );
  if (!read.ok) throw new Error(read.reason);
  const numbering = readOoxmlPart(`<w:numbering xmlns:w="${W}"/>`, {
    name: NUMBERING,
    contentType: TYPE,
  });
  if (!numbering.ok) throw new Error(numbering.reason);
  const before = read.package.partBytes.get('/[Content_Types].xml')!;
  const pkg: OoxmlPackage = {
    ...read.package,
    parts: new Map(read.package.parts).set(NUMBERING, numbering.part),
    contentTypes: {
      defaults: read.package.contentTypes.defaults,
      overrides: new Map(read.package.contentTypes.overrides).set(NUMBERING, TYPE),
    },
  };
  const store = {};
  let effects = 0;
  const stop = observeCanonicalPrimitiveJournal(
    store,
    (journal) => (effects += journal.effects.length)
  );
  let repaired: OoxmlPackage;
  try {
    repaired = runObservedStoreTransaction(
      store,
      () => withNumberingTypeBytes(pkg),
      () => true
    );
  } finally {
    stop();
  }
  expect(effects).toBe(0);
  expect(pkg.partBytes.get('/[Content_Types].xml')).toBe(before);
  const opened = readOoxmlPackage(writeOoxmlPackage(repaired));
  if (!opened.ok) throw new Error(opened.reason);
  expect(opened.package.parts.get(NUMBERING)).toBeDefined();
  expect(opened.package.contentTypes.overrides.get(NUMBERING)).toBe(TYPE);
  const xml = new TextDecoder().decode(opened.package.partBytes.get('/[Content_Types].xml'));
  expect(xml).toContain('urn:ledger');
  expect(xml).toContain('stamp="keep"');
  expect(writeOoxmlPackage(withNumberingTypeBytes(pkg))).toEqual(writeOoxmlPackage(repaired));
});

test('existing opaque numbering members survive replication and unrelated text edits', async () => {
  const read = readOoxmlPackage(
    zipDocument(body, {
      overrides: override,
      extraXml: {
        'word/numbering.xml': `<w:numbering xmlns:w="${W}" xmlns:x="urn:opaque"><w:abstractNum x:legacy="yes"><x:data>opaque</x:data></w:abstractNum><x:extension/></w:numbering>`,
      },
    })
  );
  if (!read.ok) throw new Error(read.reason);
  const expected = canonicalOoxmlFingerprint(read.package.parts.get(NUMBERING)!);
  const author = await seedReplica(read.package);
  const peer = joinReplica(author);
  try {
    const paragraphId = shownParentOf(author, findText(packageOf(author), 'Body').id, 'paragraph');
    author.doc.transact(
      () => author.registry.inline.textOf(paragraphId)!.insert(0, 'Updated '),
      'edit'
    );
    Y.applyUpdate(peer.doc, Y.encodeStateAsUpdate(author.doc), 'relay');
    for (const replica of [author, peer]) {
      const materialized = replica.materializer.rebuild();
      if (!materialized.ok) throw new Error(materialized.code);
      const saved = readOoxmlPackage(writeOoxmlPackage(packageOf(replica)));
      if (!saved.ok) throw new Error(saved.reason);
      expect(canonicalOoxmlFingerprint(saved.package.parts.get(NUMBERING)!)).toBe(expected);
    }
  } finally {
    destroyReplica(peer);
    destroyReplica(author);
  }
});

test('twenty registry lifetimes release numbering observers on a retained document', () => {
  const doc = new Y.Doc();
  const schema = packageSchemaOf(doc);
  const counts = () => {
    const listeners = (type: unknown) => {
      const value = type as { _eH: { l: unknown[] }; _dEH: { l: unknown[] } };
      return [value._eH.l.length, value._dEH.l.length];
    };
    return [
      doc._observers.get('beforeTransaction')?.size ?? 0,
      doc._observers.get('afterTransaction')?.size ?? 0,
      ...[schema.nodes, schema.attributes, schema.parts, schema.namespaces].flatMap(listeners),
    ];
  };
  const baseline = counts();
  try {
    for (let cycle = 0; cycle < 20; cycle++) {
      const registry = new DocumentRegistry(doc);
      registry.destroy();
      expect(counts()).toEqual(baseline);
    }
  } finally {
    doc.destroy();
  }
});

test('bulk loading defers live numbering restoration until derived indexes are rebuilt', async () => {
  const read = readOoxmlPackage(
    zipDocument(
      '<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>List</w:t></w:r></w:p>',
      {
        overrides: override,
        extraXml: {
          'word/numbering.xml': `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
        },
      }
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const expected = canonicalOoxmlFingerprint(read.package.parts.get(NUMBERING)!);
  const author = await seedReplica(read.package);
  try {
    const ids = author.registry.allLogicalIds().filter((id) => {
      const record = author.registry.record(id);
      return (
        record &&
        isElementRecord(record) &&
        record.namespaceUri === W &&
        ['num', 'abstractNum'].includes(record.localName)
      );
    });
    expect(ids).toHaveLength(2);
    author.registry.beginBulkLoad();
    author.registry.beginBulkLoad();
    author.doc.transact(() => {
      for (const id of ids) author.registry.schema.nodes.get(id)!.set(NODE_UNDONE_FIELD, true);
    }, 'relay');
    author.registry.endBulkLoad();
    for (const id of ids) expect(nodeRecordUndone(author.registry.schema.nodes.get(id))).toBe(true);
    author.registry.endBulkLoad();
    for (const id of ids)
      expect(nodeRecordUndone(author.registry.schema.nodes.get(id))).toBe(false);
    const peer = joinReplica(author);
    try {
      for (const replica of [author, peer]) {
        const opened = readOoxmlPackage(writeOoxmlPackage(packageOf(replica)));
        if (!opened.ok) throw new Error(opened.reason);
        expect(canonicalOoxmlFingerprint(opened.package.parts.get(NUMBERING)!)).toBe(expected);
      }
    } finally {
      destroyReplica(peer);
    }
  } finally {
    destroyReplica(author);
  }
});

test('repairing a historical numbering override preserves its opaque attributes and children', () => {
  const read = readOoxmlPackage(
    zipDocument(body, {
      overrides: `<Override PartName="${NUMBERING}" ContentType="application/legacy" xmlns:x="urn:legacy" x:stamp="original"><x:data>opaque</x:data></Override>`,
    })
  );
  if (!read.ok) throw new Error(read.reason);
  const numbering = readOoxmlPart(`<w:numbering xmlns:w="${W}"/>`, {
    name: NUMBERING,
    contentType: TYPE,
  });
  if (!numbering.ok) throw new Error(numbering.reason);
  const pkg: OoxmlPackage = {
    ...read.package,
    parts: new Map(read.package.parts).set(NUMBERING, numbering.part),
    contentTypes: {
      ...read.package.contentTypes,
      overrides: new Map(read.package.contentTypes.overrides).set(NUMBERING, TYPE),
    },
  };
  const opened = readOoxmlPackage(writeOoxmlPackage(withNumberingTypeBytes(pkg)));
  if (!opened.ok) throw new Error(opened.reason);
  expect(opened.package.parts.get(NUMBERING)).toBeDefined();
  const xml = new TextDecoder().decode(opened.package.partBytes.get('/[Content_Types].xml'));
  expect(xml).toContain('stamp="original"');
  expect(xml).toContain('<x:data>opaque</x:data>');
  expect(xml).toContain('urn:legacy');
});
