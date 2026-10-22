// Independent experiment using MIT-licensed Yjs only. Not a production adapter.
import * as Y from 'yjs';

const PREFIX = 'sciobot-owned:v1:';
type Attributes = Record<string, unknown>;
type Range = readonly [client: number, clock: number, length: number];
interface Intent {
  kind: 'insert' | 'format';
  stamp: readonly [number, number, number];
  ranges: readonly Range[];
  values: Attributes;
}
function characters(text: Y.Text) {
  const result: { client: number; clock: number }[] = [];
  for (let item = text._start; item; item = item.right) {
    if (item.deleted || !item.countable) continue;
    for (let offset = 0; offset < item.length; offset++)
      result.push({ client: item.id.client, clock: item.id.clock + offset });
  }
  return result;
}
function ranges(text: Y.Text, start: number, length: number): Range[] {
  const selected = characters(text).slice(start, start + length),
    result: [number, number, number][] = [];
  for (const item of selected) {
    const tail = result.at(-1);
    if (tail && tail[0] === item.client && tail[1] + tail[2] === item.clock) tail[2]++;
    else result.push([item.client, item.clock, 1]);
  }
  return result;
}
function stamp(doc: Y.Doc): Intent['stamp'] {
  const state = Y.decodeStateVector(Y.encodeStateVector(doc));
  return [
    [...state.values()].reduce((sum, value) => sum + value, 0) + 1,
    doc.clientID,
    state.get(doc.clientID) ?? 0,
  ];
}
const native = (attributes: Attributes) =>
  Object.fromEntries(Object.entries(attributes).filter(([key]) => !key.startsWith(PREFIX)));
function insert(
  doc: Y.Doc,
  text: Y.Text,
  index: number,
  value: string,
  attributes: Attributes = {}
) {
  doc.transact(() => {
    const order = stamp(doc),
      key = PREFIX + order.join(':');
    text.insert(index, value, { ...native(attributes), [key]: 'pending' });
    const intent: Intent = {
      kind: 'insert',
      stamp: order,
      ranges: ranges(text, index, value.length),
      values: native(attributes),
    };
    text.format(index, value.length, { [key]: JSON.stringify(intent) });
  });
}
function format(doc: Y.Doc, text: Y.Text, index: number, length: number, changes: Attributes) {
  doc.transact(() => {
    const order = stamp(doc),
      key = PREFIX + order.join(':'),
      intent: Intent = {
        kind: 'format',
        stamp: order,
        ranges: ranges(text, index, length),
        values: native(changes),
      };
    text.format(index, length, { ...intent.values, [key]: JSON.stringify(intent) });
  });
}
const compare = (a: Intent, b: Intent) =>
  a.stamp[0] - b.stamp[0] || a.stamp[1] - b.stamp[1] || a.stamp[2] - b.stamp[2];
function extended(doc: Y.Doc, intent: Intent): Intent {
  const out = intent.ranges.map((range) => [...range] as [number, number, number]);
  for (const [client, clock, length] of intent.ranges)
    for (let offset = 0; offset < length; offset++) {
      let id = Y.createID(client, clock + offset);
      for (let hops = 0; hops < 64; hops++) {
        const item = Y.getItem(doc.store, id);
        if (!(item instanceof Y.Item) || !item.redone) break;
        id = Y.createID(item.redone.client, item.redone.clock + id.clock - item.id.clock);
      }
      if (!out.some(([c, s, n]) => c === id.client && s <= id.clock && id.clock < s + n))
        out.push([id.client, id.clock, 1]);
    }
  return out.length === intent.ranges.length ? intent : { ...intent, ranges: out };
}
function normalize(doc: Y.Doc, text: Y.Text) {
  const metadata: { index: number; length: number; changes: Attributes }[] = [];
  let position = 0;
  for (const delta of text.toDelta() as { insert: string; attributes?: Attributes }[]) {
    const changes: Attributes = {};
    for (const [key, value] of Object.entries(delta.attributes ?? {}))
      if (key.startsWith(PREFIX)) {
        const original = JSON.parse(String(value)) as Intent,
          next = extended(doc, original);
        if (next !== original) changes[key] = JSON.stringify(next);
      }
    if (Object.keys(changes).length)
      metadata.push({ index: position, length: delta.insert.length, changes });
    position += delta.insert.length;
  }
  if (metadata.length)
    doc.transact(() => {
      for (const edit of metadata) text.format(edit.index, edit.length, edit.changes);
    }, 'normalization');
  const ids = characters(text),
    edits: { index: number; changes: Attributes }[] = [];
  let at = 0;
  for (const delta of text.toDelta() as { insert: string; attributes?: Attributes }[]) {
    const attributes = delta.attributes ?? {},
      intents = Object.entries(attributes)
        .filter(([key]) => key.startsWith(PREFIX))
        .map(([, value]) => JSON.parse(String(value)) as Intent)
        .sort(compare);
    for (let offset = 0; offset < delta.insert.length; offset++, at++) {
      const id = ids[at]!,
        owns = (intent: Intent) =>
          intent.ranges.some(
            ([client, clock, length]) =>
              client === id.client && clock <= id.clock && id.clock < clock + length
          );
      const baseline = intents.filter((intent) => intent.kind === 'insert' && owns(intent)).at(-1);
      if (!baseline) throw new Error(`Missing insertion owner ${id.client}:${id.clock}`);
      const target = { ...baseline.values };
      for (const intent of intents)
        if (intent.kind === 'format' && owns(intent))
          for (const [key, value] of Object.entries(intent.values)) {
            if (value === null) delete target[key];
            else target[key] = value;
          }
      const current = native(attributes),
        changes: Attributes = {};
      for (const key of new Set([...Object.keys(current), ...Object.keys(target)]))
        if (JSON.stringify(current[key]) !== JSON.stringify(target[key]))
          changes[key] = target[key] ?? null;
      if (Object.keys(changes).length) edits.push({ index: at, changes });
    }
  }
  if (!edits.length) return false;
  doc.transact(() => {
    for (const edit of edits) text.format(edit.index, 1, edit.changes);
  }, 'normalization');
  return true;
}
function reproduce(clients: readonly [number, number], receiver: 0 | 1) {
  const docs = [new Y.Doc(), new Y.Doc()];
  docs.forEach((doc, index) => (doc.clientID = clients[index]!));
  for (const doc of docs)
    doc.on('beforeObserverCalls', (transaction: Y.Transaction) => {
      if (transaction.origin === 'normalization') return;
      for (const type of transaction.changed.keys())
        if (type instanceof Y.Text) normalize(doc, type);
    });
  try {
    const alice = docs[0]!.getText('paragraph');
    insert(docs[0]!, alice, 0, 'Two', { shell: 'original' });
    Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!));
    const bob = docs[1]!.getText('paragraph'),
      undo = new Y.UndoManager(alice, { trackedOrigins: new Set(['alice']) });
    docs[0]!.transact(() => {
      insert(docs[0]!, alice, 0, 'Changed', {
        shell: 'replacement',
        insertion: 'alice',
        bold: true,
      });
      format(docs[0]!, alice, 7, 3, { shell: 'deleted', deletion: 'alice', bold: true });
    }, 'alice');
    insert(docs[1]!, bob, 3, ' Bob', { shell: 'original' });
    function settle() {
      for (let pass = 0; pass < 8; pass++) {
        const other = 1 - receiver;
        Y.applyUpdate(docs[receiver]!, Y.encodeStateAsUpdate(docs[other]!));
        Y.applyUpdate(docs[other]!, Y.encodeStateAsUpdate(docs[receiver]!));
        if (JSON.stringify(alice.toDelta()) === JSON.stringify(bob.toDelta())) return;
      }
      throw new Error('Ownership normalization did not converge');
    }
    const read = (text: Y.Text) =>
      (text.toDelta() as { insert: string; attributes?: Attributes }[]).map((span) => ({
        insert: span.insert,
        attributes: native(span.attributes ?? {}),
      }));
    settle();
    const pending = read(alice);
    undo.undo();
    settle();
    const undone = read(alice);
    undo.redo();
    settle();
    const redone = read(alice);
    undo.destroy();
    const proposed = redone
      .filter((span) => !span.attributes.deletion)
      .map((span) => span.insert)
      .join('');
    return {
      clients,
      receiver,
      pending,
      undone,
      redone,
      proposed,
      converged: JSON.stringify(read(alice)) === JSON.stringify(read(bob)),
    };
  } finally {
    docs.forEach((doc) => doc.destroy());
  }
}
const cases = (
  [
    [41, 73],
    [73, 41],
  ] as const
).flatMap((clients) => ([0, 1] as const).map((receiver) => reproduce(clients, receiver)));
process.stdout.write(
  JSON.stringify(
    {
      schemaVersion: 1,
      command: 'bun packages/sciobot/test/diagnostics/yjs-owned-format-proof.ts',
      cases,
    },
    null,
    2
  ) + '\n'
);
if (cases.some((item) => item.proposed !== 'Changed Bob' || !item.converged)) process.exitCode = 1;
