import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  runWithRevisionIdReservations,
  REVISION_ID_RESERVATION_LIMITS,
  TreeDocumentStore,
  TreePackageStore,
  type OoxmlPart,
  type RevisionIdReservation,
  type TreeDocOp,
} from '../index.ts';
import {
  ACTOR_ID_STRIPE,
  actorStripe,
  MAX_DECIMAL_ID,
  runWithTransactionActor,
} from '../package/actor-scoped-ids.ts';
import { mintFragmentUniqueIds } from '../store/clipboard-fragment-unique-ids.ts';
import { nextRevisionId } from '../store/tree-op-revision-ids.ts';
import { ADA, paragraphId, part } from './tracked-edit-fixture.ts';

const paragraph = '<w:p><w:r><w:t>Text</w:t></w:r></w:p>';
const source = () => part(paragraph);
const reservations = (ids: string[], name = '/word/document.xml'): RevisionIdReservation[] => [
  { partName: name, ids },
];
const nativeIds = (value: OoxmlPart) => [
  ...new Set(
    collectRevisionSites(value).map(
      (site) => site.node.attributes.find((attribute) => attribute.localName === 'id')!.value
    )
  ),
];

function apply(value: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(value, op);
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

test('dense allocation reserves retired IDs and retains native live IDs without document writes', () => {
  const value = part(
      '<w:p><w:ins w:id="2" w:author="Peer"><w:r><w:t>Live</w:t></w:r></w:ins></w:p>'
    ),
    before = canonicalOoxmlFingerprint(value);
  runWithRevisionIdReservations(reservations(['+0004', '00000000000000000000000005']), () => {
    const mint = nextRevisionId(value);
    expect([mint(), mint()]).toEqual(['6', '7']);
    expect(nextRevisionId(value)()).toBe('6');
    expect(nextRevisionId({ ...value, name: '/word/header1.xml' })()).toBe('3');
  });
  expect(nextRevisionId(value)()).toBe('3');
  expect(canonicalOoxmlFingerprint(value)).toBe(before);
});

test('dense revision allocation wraps at the Word ceiling and skips retained and live spellings', () => {
  const value = part(
    '<w:p><w:ins w:id="+0001" w:author="Peer"><w:r><w:t>Live</w:t></w:r></w:ins></w:p>'
  );
  const mint = runWithRevisionIdReservations(reservations(['0', '2', String(MAX_DECIMAL_ID)]), () =>
    nextRevisionId(value)
  );
  expect([mint(), mint(), mint()]).toEqual(['3', '4', '5']);
});

test('invalid decimal strings cannot raise allocation outside Word bounds', () => {
  runWithRevisionIdReservations(
    reservations(['-1', '1e22', '99999999999999999', '2147483648', 'not-an-id']),
    () => {
      expect(nextRevisionId(source())()).toBe('0');
    }
  );
});

for (const actor of ['alice', 'bob'])
  test(`actor ${actor} retains its native stripe while skipping retained and live IDs`, () => {
    const stripe = actorStripe(actor),
      value = part(
        `<w:p><w:ins w:id="${stripe}" w:author="Peer"><w:r><w:t>Live</w:t></w:r></w:ins></w:p>`
      );
    runWithRevisionIdReservations(reservations([String(stripe + ACTOR_ID_STRIPE)]), () => {
      const explicit = nextRevisionId(value, actor),
        implicit = runWithTransactionActor(actor, () => nextRevisionId(value));
      expect([explicit(), explicit()]).toEqual([
        String(stripe + 2 * ACTOR_ID_STRIPE),
        String(stripe + 3 * ACTOR_ID_STRIPE),
      ]);
      expect(implicit()).toBe(String(stripe + 2 * ACTOR_ID_STRIPE));
    });
  });

test('a fully reserved actor stripe refuses instead of leaving the stripe or reusing history', () => {
  const stripe = actorStripe('alice'),
    ids = Array.from({ length: 32768 }, (_, index) => String(stripe + index * ACTOR_ID_STRIPE));
  runWithRevisionIdReservations(reservations(ids), () => {
    expect(() => nextRevisionId(source(), 'alice')()).toThrow('no free decimal id');
  });
});

test('nested scopes replace and restore owned copies after success, refusal, or throw', () => {
  const ids = ['8'];
  runWithRevisionIdReservations(reservations(ids), () => {
    ids.push('90');
    expect(nextRevisionId(source())()).toBe('9');
    runWithRevisionIdReservations([], () => expect(nextRevisionId(source())()).toBe('0'));
    expect(() =>
      runWithRevisionIdReservations(reservations(['20']), () => {
        expect(nextRevisionId(source())()).toBe('21');
        throw new Error('stop');
      })
    ).toThrow('stop');
    expect(() =>
      runWithRevisionIdReservations(reservations([42 as unknown as string]), () => {})
    ).toThrow('Invalid reserved revision ID');
    expect(nextRevisionId(source())()).toBe('9');
  });
  expect(nextRevisionId(source())()).toBe('0');
});

test('invalid reservation structures and bounded excess refuse before the callback', () => {
  let called = false;
  for (const invalid of [
    null,
    [{ partName: 'relative.xml', ids: [] }],
    [{ partName: '/word/../other.xml', ids: [] }],
    [{ partName: '/word/document.xml', ids: null }],
    reservations(['0'.repeat(REVISION_ID_RESERVATION_LIMITS.maxStringLength + 1)]),
    reservations(Array(REVISION_ID_RESERVATION_LIMITS.maxIds + 1).fill('1')),
    Array(REVISION_ID_RESERVATION_LIMITS.maxParts + 1).fill({
      partName: '/word/document.xml',
      ids: [],
    }),
  ]) {
    expect(() =>
      runWithRevisionIdReservations(invalid as RevisionIdReservation[], () => {
        called = true;
      })
    ).toThrow();
  }
  expect(called).toBe(false);
  expect(nextRevisionId(source())()).toBe('0');
});

for (const kind of ['text', 'format', 'row', 'attribution'] as const)
  test(`native ${kind} allocation uses retained reservations`, () => {
    const value = part(
        paragraph +
          '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
      ),
      body = value.root.children[0]!,
      table =
        body.kind === 'textValue' ? undefined : body.children.find((node) => node.kind === 'table'),
      row =
        table?.kind === 'table'
          ? table.children.find((node) => node.kind === 'tableRow')
          : undefined;
    if (!table || !row) throw new Error('Missing table');
    runWithRevisionIdReservations(reservations(['8']), () => {
      let result: OoxmlPart;
      if (kind === 'text')
        result = apply(value, {
          op: 'insertText',
          paragraphId: paragraphId(value),
          offset: 1,
          text: 'New',
          revision: ADA,
        });
      else if (kind === 'format')
        result = apply(value, {
          op: 'setRunProperties',
          paragraphId: paragraphId(value),
          start: 0,
          end: 4,
          properties: [{ localName: 'b', attributes: {} }],
          revision: ADA,
        });
      else if (kind === 'row')
        result = apply(value, {
          op: 'insertTableRow',
          tableId: table.id,
          rowId: row.id,
          where: 'below',
          revision: ADA,
        });
      else {
        const pending = apply(value, {
          op: 'insertText',
          paragraphId: paragraphId(value),
          offset: 1,
          text: 'New',
          revision: ADA,
        });
        result = apply(pending, {
          op: 'setRevisionAttribution',
          revision: ADA,
          sites: collectRevisionSites(pending).map((site) => ({
            nodeId: site.node.id,
            renumber: 'independent',
          })),
        });
      }
      expect(nativeIds(result)).toEqual([kind === 'attribution' ? '10' : '9']);
    });
  });

test('package transactions re-read host reservations and retain one native history unit', () => {
  const value = source(),
    pkg = new TreeDocumentStore(value).package;
  let ids = ['8'],
    reads = 0;
  const store = new TreePackageStore(pkg, value, {
    reservedRevisionIds: (current) => {
      expect(current.parts.has(value.name)).toBe(true);
      reads++;
      return reservations(ids);
    },
  });
  const insert = () =>
    store.transact({ kind: 'body' }, (ctx) => {
      ctx.apply({
        op: 'insertText',
        paragraphId: paragraphId(value),
        offset: 1,
        text: 'New',
        revision: ADA,
      });
    });
  expect(insert().ok).toBe(true);
  expect(nativeIds(store.partFor({ kind: 'body' })!)).toEqual(['9']);
  expect(
    store.transact({ kind: 'body' }, (ctx) => {
      ctx.apply({ op: 'acceptAllRevisions' });
    }).ok
  ).toBe(true);
  ids = ['8', '9', '20'];
  expect(insert().ok).toBe(true);
  expect(nativeIds(store.partFor({ kind: 'body' })!)).toEqual(['21']);
  expect(reads).toBe(3);
  expect(store.undo()).not.toBeNull();
  expect(nativeIds(store.partFor({ kind: 'body' })!)).toEqual([]);
  expect(store.undo()).not.toBeNull();
  expect(nativeIds(store.partFor({ kind: 'body' })!)).toEqual(['9']);
});

for (const actor of [undefined, 'alice'])
  test(`clipboard native revisions reserve retained IDs including Word ceiling; actor ${actor ?? 'solo'}`, () => {
    const value = source(),
      pkg = new TreeDocumentStore(value).package,
      fragment = part(
        '<w:p><w:ins w:id="1" w:author="Peer"><w:r><w:t>Paste</w:t></w:r></w:ins></w:p>'
      ),
      ids = actor ? [String(actorStripe(actor))] : ['0', String(MAX_DECIMAL_ID)];
    runWithRevisionIdReservations(reservations(ids), () =>
      runWithTransactionActor(actor, () => {
        const result = mintFragmentUniqueIds({
          pkg,
          ownerPart: value,
          notesParts: [],
          travelling: fragment.root.children,
          fragmentBookmarkIds: [],
          hasRevision: true,
          hasDocPr: false,
        });
        expect(result?.revisionIds.get('1')).toBe(
          actor ? String(actorStripe(actor) + ACTOR_ID_STRIPE) : '1'
        );
      })
    );
  });
