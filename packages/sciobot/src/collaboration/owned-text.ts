import * as Y from 'yjs';
import { copyBoundedJson } from '../data/json';
import {
  compareIntents,
  decodeIntent,
  decodeLineage,
  detachedAttributes,
  encodeIntent,
  nativeAttributes,
  OWNED_TEXT_OPERATIONS,
  OWNED_TEXT_LINEAGE,
  OWNED_TEXT_PREFIX,
  OWNED_TEXT_PROTOCOL,
  OwnedTextError,
  type Attributes,
  type AttributeIntent,
} from './ownership-codec';
import {
  observedRanges,
  ownershipCheck,
  recordRestoredLineage,
  visibleSpans,
} from './ownership-positions';

export { OWNED_TEXT_PROTOCOL, OwnedTextError } from './ownership-codec';
export interface OwnedTextOptions {
  readonly onError?: (error: Error) => void;
}
export interface OwnedTextHandle {
  readonly protocol: typeof OWNED_TEXT_PROTOCOL;
  assertHealthy(): void;
  dispose(): void;
}
interface State {
  readonly doc: Y.Doc;
  readonly options: OwnedTextOptions;
  error: Error | null;
  repairing: boolean;
  /** Texts changed while the room held updates with missing dependencies. */
  readonly deferred: Set<Y.Text>;
}
const attached = new WeakMap<Y.Doc, State>();
const REPAIR = Symbol('sciobot-owned-text-normalization');
const original = {
  insert: Y.Text.prototype.insert,
  insertEmbed: Y.Text.prototype.insertEmbed,
  format: Y.Text.prototype.format,
  applyDelta: Y.Text.prototype.applyDelta,
};
let patched = false;
const lengthOf = (value: unknown) => (typeof value === 'string' ? value.length : 1);
function validRange(text: Y.Text, index: number, length: number) {
  if (
    !Number.isSafeInteger(index) ||
    index < 0 ||
    !Number.isSafeInteger(length) ||
    length < 0 ||
    index + length > text.length
  )
    throw new OwnedTextError('Invalid shared-text UTF-16 range.');
}
function check(state: State) {
  if (state.error) throw state.error;
}
function nextIntent(
  doc: Y.Doc,
  kind: AttributeIntent['kind'],
  ranges: AttributeIntent['ranges'],
  values: Attributes
) {
  const vector = Y.decodeStateVector(Y.encodeStateVector(doc)),
    clock = vector.get(doc.clientID) ?? 0,
    stamp = [
      1 + [...vector.values()].reduce((sum, value) => sum + value, 0),
      doc.clientID,
      clock,
    ] as const;
  return {
    key: `${doc.clientID}:${clock}`,
    intent: { kind, stamp, ranges, values } satisfies AttributeIntent,
  };
}
function record(
  state: State,
  text: Y.Text,
  index: number,
  length: number,
  kind: AttributeIntent['kind'],
  values: Attributes
) {
  publish(state, text, index, length, prepare(state, text, index, length, kind, values));
}
function prepare(
  state: State,
  text: Y.Text,
  index: number,
  length: number,
  kind: AttributeIntent['kind'],
  values: Attributes
) {
  const ranges = observedRanges(text, index, length);
  if (ranges.length > 100_000)
    throw new OwnedTextError('Owned attribute range count exceeds its limit.');
  const item = nextIntent(state.doc, kind, ranges, values);
  return { ...item, data: encodeIntent(item.intent) };
}
function publish(
  state: State,
  text: Y.Text,
  index: number,
  length: number,
  item: ReturnType<typeof prepare>
) {
  const map = state.doc.getMap<string>(OWNED_TEXT_OPERATIONS);
  const version = map.get('protocol');
  if (version !== undefined && version !== '1')
    throw new OwnedTextError('Unsupported owned-text protocol.');
  if (version === undefined) map.set('protocol', '1');
  if (map.has(item.key)) throw new OwnedTextError('Owned attribute identity was reused.');
  map.set(item.key, item.data);
  original.format.call(text, index, length, { [OWNED_TEXT_PREFIX + item.key]: 1 });
}
function captureInsert(state: State, text: Y.Text, index: number, length: number) {
  let position = 0;
  const runs: { index: number; length: number; attributes: Attributes }[] = [];
  for (const delta of text.toDelta() as { insert: unknown; attributes?: Attributes }[]) {
    const count = lengthOf(delta.insert),
      start = Math.max(index, position),
      end = Math.min(index + length, position + count);
    if (start < end)
      runs.push({
        index: start,
        length: end - start,
        attributes: detachedAttributes(delta.attributes ?? {}),
      });
    position += count;
  }
  for (const run of runs) record(state, text, run.index, run.length, 'insert', run.attributes);
}
function install() {
  if (patched) return;
  patched = true;
  Y.Text.prototype.insert = function (index, value, attributes) {
    const state = this.doc ? attached.get(this.doc) : undefined;
    if (!state || !value.length) return original.insert.call(this, index, value, attributes);
    check(state);
    if (!Number.isSafeInteger(index) || index < 0)
      throw new OwnedTextError('Invalid shared-text insertion index.');
    const at = Math.min(index, this.length),
      attrs = attributes === undefined ? undefined : detachedAttributes(attributes);
    state.doc.transact(() => {
      original.insert.call(this, at, value, attrs);
      captureInsert(state, this, at, value.length);
    });
  };
  Y.Text.prototype.insertEmbed = function (index, value, attributes) {
    const state = this.doc ? attached.get(this.doc) : undefined;
    if (!state) return original.insertEmbed.call(this, index, value, attributes);
    check(state);
    if (!Number.isSafeInteger(index) || index < 0)
      throw new OwnedTextError('Invalid shared-text embed index.');
    const at = Math.min(index, this.length),
      attrs = attributes === undefined ? undefined : detachedAttributes(attributes);
    state.doc.transact(() => {
      original.insertEmbed.call(this, at, value, attrs);
      captureInsert(state, this, at, 1);
    });
  };
  Y.Text.prototype.format = function (index, length, attributes) {
    const state = this.doc ? attached.get(this.doc) : undefined;
    if (!state) return original.format.call(this, index, length, attributes);
    check(state);
    validRange(this, index, length);
    if (!length) return;
    const values = detachedAttributes(attributes);
    if (!Object.keys(values).length) return;
    const item = prepare(state, this, index, length, 'format', values);
    state.doc.transact(() => {
      original.format.call(this, index, length, values);
      publish(state, this, index, length, item);
    });
  };
  Y.Text.prototype.applyDelta = function (delta, options) {
    const state = this.doc ? attached.get(this.doc) : undefined;
    if (!state) return original.applyDelta.call(this, delta, options);
    check(state);
    const copy = copyBoundedJson(delta, {
      maxBytes: 16 * 1024 * 1024,
      maxNodes: 1_000_000,
      maxDepth: 32,
    });
    if (!copy.ok || !Array.isArray(copy.value))
      throw new OwnedTextError('Invalid bounded shared-text delta.');
    const operations = copy.value as typeof delta;
    let cursor = 0,
      size = this.length;
    for (const op of operations) {
      if (
        !op ||
        typeof op !== 'object' ||
        Array.isArray(op) ||
        Object.keys(op).some(
          (key) => !['insert', 'retain', 'delete', 'attributes'].includes(key)
        ) ||
        ['insert', 'retain', 'delete'].filter((key) => Object.hasOwn(op, key)).length !== 1
      )
        throw new OwnedTextError('Invalid shared-text delta operation.');
      if (op.attributes !== undefined) detachedAttributes(op.attributes);
      if (op.insert !== undefined) {
        const length = lengthOf(op.insert);
        cursor += length;
        size += length;
      } else {
        const length = op.retain ?? op.delete;
        if (!Number.isSafeInteger(length) || length! < 0 || cursor + length! > size)
          throw new OwnedTextError('Invalid shared-text delta range.');
        if (op.retain !== undefined) cursor += length!;
        else size -= length!;
      }
    }
    state.doc.transact(() => {
      let at = 0;
      for (const [index, op] of operations.entries()) {
        if (op.insert !== undefined) {
          const value =
            typeof op.insert === 'string' &&
            options?.sanitize === false &&
            index === delta.length - 1 &&
            at === this.length &&
            op.insert.endsWith('\n')
              ? op.insert.slice(0, -1)
              : op.insert;
          if (typeof value === 'string') {
            if (value.length) this.insert(at, value, op.attributes);
            at += value.length;
          } else {
            this.insertEmbed(at, value, op.attributes);
            at++;
          }
        } else if (op.retain !== undefined) {
          if (op.attributes) this.format(at, op.retain, op.attributes);
          at += op.retain;
        } else if (op.delete !== undefined) this.delete(at, op.delete);
      }
    });
  };
}
function repair(state: State, text: Y.Text) {
  if (text.length === 0) return;
  const map = state.doc.share.get(OWNED_TEXT_OPERATIONS);
  if (!(map instanceof Y.Map) || map.get('protocol') !== '1')
    throw new OwnedTextError('Shared text has no owned-attribute protocol.');
  const parsed = new Map<string, AttributeIntent>(),
    spans = visibleSpans(text),
    owns = ownershipCheck(state.doc),
    edits: {
      index: number;
      length: number;
      changes: Record<string, unknown>;
      signature: string;
    }[] = [];
  let position = 0,
    spanIndex = 0;
  for (const delta of text.toDelta() as { insert: unknown; attributes?: Attributes }[]) {
    const attributes = delta.attributes ?? {},
      intents: { key: string; intent: AttributeIntent }[] = [];
    for (const key of Object.keys(attributes))
      if (key.startsWith(OWNED_TEXT_PREFIX)) {
        if (attributes[key] !== 1)
          throw new OwnedTextError('Invalid owned attribute activity marker.');
        let intent = parsed.get(key);
        if (!intent) {
          intent = decodeIntent(map.get(key.slice(OWNED_TEXT_PREFIX.length)));
          parsed.set(key, intent);
        }
        intents.push({ key, intent });
      }
    intents.sort((a, b) => compareIntents(a.intent, b.intent));
    const current = nativeAttributes(attributes);
    for (let offset = 0; offset < lengthOf(delta.insert); offset++, position++) {
      while (
        spanIndex < spans.length &&
        position >= spans[spanIndex]!.index + spans[spanIndex]!.length
      )
        spanIndex++;
      const span = spans[spanIndex];
      if (!span) throw new OwnedTextError('Shared text identity does not match its content.');
      const clock = span.clock + position - span.index,
        owned = intents.filter((item) => owns(item.intent.ranges, span.client, clock)),
        baseline = owned.filter((item) => item.intent.kind === 'insert').at(-1);
      if (!baseline) throw new OwnedTextError('A shared character has no insertion owner.');
      const target: Record<string, unknown> = Object.assign(
        Object.create(null),
        baseline.intent.values
      );
      for (const { intent } of owned)
        if (intent.kind === 'format')
          for (const [key, value] of Object.entries(intent.values)) {
            if (value === null) delete target[key];
            else target[key] = value;
          }
      const changes: Record<string, unknown> = Object.create(null);
      for (const key of new Set([...Object.keys(current), ...Object.keys(target)]))
        if (JSON.stringify(current[key]) !== JSON.stringify(target[key]))
          changes[key] = target[key] ?? null;
      for (const item of intents) if (!owned.includes(item)) changes[item.key] = null;
      if (!Object.keys(changes).length) continue;
      const signature = JSON.stringify(changes),
        last = edits.at(-1);
      if (last && last.index + last.length === position && last.signature === signature)
        last.length++;
      else edits.push({ index: position, length: 1, changes, signature });
    }
  }
  if (edits.length)
    state.doc.transact(() => {
      for (const edit of edits) original.format.call(text, edit.index, edit.length, edit.changes);
    }, REPAIR);
}
function changedTexts(transaction: Y.Transaction) {
  const output = new Set<Y.Text>(),
    seen = new Set<unknown>();
  const visit = (value: unknown) => {
    if (seen.has(value)) return;
    seen.add(value);
    if (value instanceof Y.Text) {
      output.add(value);
      return;
    }
    if (value instanceof Y.Map) for (const child of value.values()) visit(child);
    else if (value instanceof Y.Array) for (const child of value.toArray()) visit(child);
  };
  for (const [type, keys] of transaction.changed) {
    if (type instanceof Y.Text) output.add(type);
    else if (type instanceof Y.Map) {
      for (const key of keys) if (key !== null) visit(type.get(key));
    } else if (type instanceof Y.Array) visit(type);
  }
  return output;
}
function validateChangedLedgers(doc: Y.Doc, transaction: Y.Transaction) {
  const operations = doc.getMap(OWNED_TEXT_OPERATIONS),
    lineage = doc.getMap(OWNED_TEXT_LINEAGE);
  const changed = [...transaction.changed].find(
    ([type]) => type instanceof Y.Map && type === operations
  )?.[1];
  if (changed) {
    if (operations.get('protocol') !== '1')
      throw new OwnedTextError('Unsupported owned-text protocol.');
    for (const key of changed) {
      if (key === null || key === 'protocol') continue;
      const intent = decodeIntent(operations.get(key));
      if (key !== `${intent.stamp[1]}:${intent.stamp[2]}`)
        throw new OwnedTextError('Owned attribute address does not match its record.');
    }
  }
  const changedLineage = [...transaction.changed].find(
    ([type]) => type instanceof Y.Map && type === lineage
  )?.[1];
  for (const key of changedLineage ?? []) {
    if (key === null || !/^\d+:\d+$/.test(key))
      throw new OwnedTextError('Invalid restored character address.');
    const [client, clock] = key.split(':').map(Number),
      value = decodeLineage(lineage.get(key));
    if (
      !Number.isSafeInteger(client) ||
      !Number.isSafeInteger(clock) ||
      clock! + value.length > Number.MAX_SAFE_INTEGER
    )
      throw new OwnedTextError('Invalid restored character clock.');
  }
}
/**
 * Attach before native shared-text authoring or a room snapshot is applied.
 * Every participant must use this protocol. The host supplies the pinned MIT Yjs peer.
 * No evaluated transport implementation is imported or modified.
 */
export function attachOwnedTextAttributes(
  doc: Y.Doc,
  options: OwnedTextOptions = {}
): OwnedTextHandle {
  if (!(doc instanceof Y.Doc) || doc.isDestroyed)
    throw new OwnedTextError('The room must use the same pinned Yjs instance.');
  if (attached.has(doc))
    throw new OwnedTextError('Owned attributes are already attached to this room.');
  const operations = doc.getMap(OWNED_TEXT_OPERATIONS);
  const protocol = operations.get('protocol');
  if (protocol !== undefined && protocol !== '1')
    throw new OwnedTextError('Unsupported owned-text protocol.');
  if (protocol === undefined && [...doc.store.clients.values()].some((items) => items.length))
    throw new OwnedTextError('Attach owned attributes before legacy room content is authored.');
  doc.getMap(OWNED_TEXT_LINEAGE);
  install();
  const state: State = { doc, options, error: null, repairing: false, deferred: new Set() };
  attached.set(doc, state);
  const before = (transaction: Y.Transaction) => {
    if (transaction.origin === REPAIR || state.repairing) return;
    try {
      check(state);
      state.repairing = true;
      recordRestoredLineage(doc, transaction, REPAIR);
      validateChangedLedgers(doc, transaction);
      const texts = changedTexts(transaction);
      // A partial delivery can apply a peer's repair before that peer's intents arrive.
      // Repairing with that partial knowledge would undo it and the peers would loop.
      if (doc.store.pendingStructs || doc.store.pendingDs) {
        for (const text of texts) state.deferred.add(text);
        return;
      }
      for (const text of state.deferred) texts.add(text);
      state.deferred.clear();
      for (const text of texts) if (!text._item?.deleted && text.doc === doc) repair(state, text);
    } catch (error) {
      state.error = error instanceof Error ? error : new OwnedTextError(String(error));
      state.options.onError?.(state.error);
      throw state.error;
    } finally {
      state.repairing = false;
    }
  };
  doc.on('beforeObserverCalls', before);
  let disposed = false;
  return Object.freeze({
    protocol: OWNED_TEXT_PROTOCOL,
    assertHealthy() {
      check(state);
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      doc.off('beforeObserverCalls', before);
      state.deferred.clear();
      attached.delete(doc);
    },
  });
}
