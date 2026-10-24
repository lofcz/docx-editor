import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { nativeBirthFieldCreation } from './review-birth-field-retention';
import { attachOwnedReviewFields } from './review-fields';
import {
  REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
  decodeFieldTarget,
} from './review-field-codec';

const META = 'docx-package-meta-v1',
  NODES = 'docx-package-nodes-v1',
  ATTRS = 'docx-package-attributes-v1';
const key = (id: string, ns: string, name: string) => [id, ns, name].join('\u001f');
function fixture(repair = 14) {
  const doc = new Y.Doc();
  const meta = doc.getMap(META),
    nodes = doc.getMap<Y.Map<unknown>>(NODES),
    attrs = doc.getMap<string>(ATTRS);
  for (const [name, value] of Object.entries({
    protocolVersion: 1,
    sharedSchemaVersion: 8,
    repairVersion: repair,
    canonicalModelVersion: 1,
  }))
    meta.set(name, value);
  const fields = attachOwnedReviewFields(doc);
  fields.enable();
  return {
    doc,
    nodes,
    attrs,
    fields,
    destroy() {
      fields.dispose();
      doc.destroy();
    },
  };
}
function node(localName: string) {
  const value = new Y.Map<unknown>();
  value.set('s', ['generic', 'w', localName, ''].join('\u001f'));
  return value;
}

test('native birth constructor fields stay outside conditional scalar ownership; later writes keep their baseline', () => {
  const f = fixture();
  const geometry = key('width', 'w', 'w'),
    identity = key('paragraph', 'w14', 'paraId');
  try {
    f.fields.capture(
      () =>
        f.doc.transact(() => {
          f.nodes.set('width', node('tcW'));
          f.nodes.set('paragraph', node('p'));
          f.attrs.set(geometry, '2000');
          f.attrs.set(identity, '1234ABCD');
        }),
      8
    );
    expect(f.doc.getMap(REVIEW_FIELD_TARGETS).size).toBe(0);
    expect(f.doc.getMap(REVIEW_FIELD_INTENTS).size).toBe(0);
    const history = new Y.UndoManager(f.attrs, { captureTimeout: 0 });
    try {
      f.fields.capture(() => f.doc.transact(() => f.attrs.set(geometry, '3000')), 1);
      const targets = f.doc.getMap(REVIEW_FIELD_TARGETS);
      expect(targets.get('protocol')).toBe('1');
      const entries = [...targets.entries()].filter(([address]) => address !== 'protocol');
      expect(entries.length).toBe(1);
      const [address, raw] = entries[0]!;
      expect(decodeFieldTarget(address, raw).baseline).toBe('2000');
      expect(f.attrs.get(geometry)).toBe('3000');
      expect(history.undo()).not.toBeNull();
      expect(f.attrs.get(geometry)).toBe('2000');
      expect(f.attrs.get(identity)).toBe('1234ABCD');
      expect(history.redo()).not.toBeNull();
      expect(f.attrs.get(geometry)).toBe('3000');
    } finally {
      history.destroy();
    }
  } finally {
    f.destroy();
  }
});

for (const repair of [13, 15])
  test(`constructor fields remain conditional outside exact repair ${repair}`, () => {
    const f = fixture(repair);
    try {
      f.fields.capture(
        () =>
          f.doc.transact(() => {
            f.nodes.set('width', node('tcW'));
            f.attrs.set(key('width', 'w', 'w'), '2000');
          }),
        8
      );
      expect(f.doc.getMap(REVIEW_FIELD_TARGETS).size).toBeGreaterThan(0);
    } finally {
      f.destroy();
    }
  });

test('native birth exemption excludes old nodes, unrelated fields, other roots, and sequences', () => {
  const f = fixture();
  f.nodes.set('old', node('tcW'));
  let assertions = 0;
  const check = (tx: Y.Transaction) => {
    expect(nativeBirthFieldCreation(f.doc, tx, { root: ATTRS, key: key('new', 'w', 'w') })).toBe(
      true
    );
    for (const target of [
      { root: ATTRS, key: key('old', 'w', 'w') },
      { root: ATTRS, key: key('other', 'w', 'val') },
      { root: 'other', key: key('new', 'w', 'w') },
      { root: ATTRS, sequence: true as const, key: key('new', 'w', 'w') },
      { root: ATTRS, key: 'malformed' },
      { root: ATTRS, key: key('missing', 'w14', 'paraId') },
    ])
      expect(nativeBirthFieldCreation(f.doc, tx, target)).toBe(false);
    assertions++;
  };
  try {
    f.doc.once('afterTransaction', check);
    f.doc.transact(() => {
      f.nodes.set('new', node('tcW'));
      f.nodes.set('other', node('color'));
    });
    expect(assertions).toBe(1);
  } finally {
    f.destroy();
  }
});

test('wire inspection does not create absent native roots', () => {
  const doc = new Y.Doc();
  try {
    doc.once('afterTransaction', (tx: Y.Transaction) => {
      expect(
        nativeBirthFieldCreation(doc, tx, { root: ATTRS, key: key('new', 'w14', 'paraId') })
      ).toBe(false);
      expect(doc.share.size).toBe(0);
    });
    doc.transact(() => {});
    expect(doc.share.size).toBe(0);
  } finally {
    doc.destroy();
  }
});
