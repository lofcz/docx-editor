// Application-owned review string fields over pinned MIT Yjs. No provider schema is read.
import * as Y from 'yjs';
import { OWNED_TEXT_OPERATIONS, OWNED_TEXT_LINEAGE } from './ownership-codec';
import { REVIEW_FLAG_TARGETS, REVIEW_FLAG_INTENTS } from './review-flag-codec';
import { flagMap } from './review-flag-items';
import { recordRestoredLineage } from './ownership-positions';
import {
  reviewArray,
  removedStringItems,
  normalizeReviewArray,
  originalReviewSequenceTarget,
} from './review-sequences';
import type { OwnedTextOptions } from './owned-text';
import {
  REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
  MAX_REVIEW_FIELD_TARGETS,
  MAX_REVIEW_FIELD_INTENTS,
  ReviewFieldError,
  fieldAddress,
  encodeFieldRecord,
  decodeFieldTarget,
  decodeFieldIntent,
  newerFieldIntent,
  type ReviewFieldLocation,
  type ReviewFieldTarget,
  type ReviewFieldIntent,
} from './review-field-codec';
export { REVIEW_FIELD_TARGETS, REVIEW_FIELD_INTENTS } from './review-field-codec';
export interface OwnedReviewFieldOptions extends OwnedTextOptions {
  /** Hosts may lower the fixed record limits. Raising them is refused. */
  readonly limits?: { readonly targets?: number; readonly intents?: number };
}
interface Reservation {
  remaining: number;
  readonly captured: Map<string, string>;
}
interface State {
  readonly doc: Y.Doc;
  readonly targets: Y.Map<string>;
  readonly intents: Y.Map<string>;
  readonly origins: Set<unknown>;
  readonly known: WeakMap<Y.Map<unknown>, Map<string, string>>;
  readonly transactions: WeakMap<Y.Transaction, Map<string, string>>;
  readonly sequencePresence: Map<string, boolean>;
  readonly reservations: Reservation[];
  readonly limits: { readonly targets: number; readonly intents: number };
  enabled: boolean;
  capturing: boolean;
  captureOrigin: { readonly value: unknown } | null;
  normalizing: boolean;
  deferred: boolean;
  disposed: boolean;
  error: Error | null;
}
const NORMALIZE = Symbol('sciobot-review-field-normalization');
const attached = new WeakMap<Y.Doc, State>();
const excluded = new Set([
  REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
  REVIEW_FLAG_TARGETS,
  REVIEW_FLAG_INTENTS,
  OWNED_TEXT_OPERATIONS,
  OWNED_TEXT_LINEAGE,
]);
let installed = false;
let set: typeof Y.Map.prototype.set, remove: typeof Y.Map.prototype.delete;
let removeItems: typeof Y.Array.prototype.delete;
let insertItems: typeof Y.Array.prototype.insert;
function healthy(state: State) {
  if (state.error) throw state.error;
  if (state.disposed || state.doc.isDestroyed)
    throw new ReviewFieldError('The owned review field scope is no longer active.');
}
function rootName(state: State, map: Y.Map<unknown> | Y.Array<unknown>): string | null {
  let type: unknown = map;
  for (let depth = 0; depth < 256; depth++) {
    const item = (type as { readonly _item: Y.Item | null })._item;
    if (!item) {
      for (const [name, root] of state.doc.share) if ((root as unknown) === type) return name;
      return null;
    }
    type = item.parent;
  }
  throw new ReviewFieldError('Review field shared-map ancestry exceeds its limit.');
}
function location(state: State, map: Y.Map<unknown>, key: string): ReviewFieldLocation | null {
  const root = rootName(state, map);
  if (root === null || excluded.has(root)) return null;
  if (typeof key !== 'string' || !key.length || key.length > 65536) {
    if (state.capturing) throw new ReviewFieldError('Invalid bounded review field key.');
    return null;
  }
  const known = state.known.get(map)?.get(key);
  if (known) return decodeFieldTarget(known, state.targets.get(known));
  return map._item
    ? { client: map._item.id.client, clock: map._item.id.clock, key }
    : { root, key };
}
function eligible(
  state: State,
  map: Y.Map<unknown>,
  key: string,
  value: string | null
): ReviewFieldLocation | null {
  if (!state.enabled || state.normalizing || state.disposed) return null;
  const transaction = state.doc._transaction;
  if (
    transaction &&
    (!transaction.local ||
      transaction.origin === NORMALIZE ||
      transaction.origin instanceof Y.UndoManager)
  )
    return null;
  if (state.capturing && state.captureOrigin && state.captureOrigin.value !== transaction?.origin)
    return null;
  const target = location(state, map, key);
  if (!target) return null;
  if (
    map._item &&
    transaction &&
    map._item.id.clock >= (transaction.beforeState.get(map._item.id.client) ?? 0)
  )
    return null;
  if (
    !state.capturing &&
    target.root !== undefined &&
    !map.has(key) &&
    !state.targets.has(fieldAddress(target))
  )
    return null;
  const previous = map.get(key);
  if (previous !== undefined && typeof previous !== 'string') return null;
  if (value === null && previous === undefined && !state.targets.has(fieldAddress(target)))
    return null;
  return target;
}
function captureMap(state: State): Map<string, string> {
  const reserved = state.reservations.at(-1);
  if (reserved) return reserved.captured;
  const transaction = state.doc._transaction;
  if (!transaction) throw new ReviewFieldError('Owned review fields need a shared transaction.');
  let captured = state.transactions.get(transaction);
  if (!captured) {
    captured = new Map();
    state.transactions.set(transaction, captured);
  }
  return captured;
}
function record(
  state: State,
  map: Y.Map<unknown> | null,
  target: ReviewFieldLocation,
  value: string | null,
  sequenceBaseline?: {
    value: string;
    position: ReviewFieldTarget['position'];
    initiallyPresent: boolean;
  }
) {
  healthy(state);
  if (state.doc.store.pendingStructs || state.doc.store.pendingDs)
    throw new ReviewFieldError('Review fields must wait for complete room updates.');
  const key = fieldAddress(target),
    captured = captureMap(state),
    previousIntent = captured.get(key),
    reservation = state.reservations.at(-1);
  if (!previousIntent && reservation && !reservation.remaining)
    throw new ReviewFieldError('The review command exceeded its planned field count.');
  if (
    (!state.targets.has(key) &&
      state.targets.size - Number(state.targets.has('protocol')) >= state.limits.targets) ||
    (!previousIntent && state.intents._map.size >= state.limits.intents)
  )
    throw new ReviewFieldError('Review field ownership exceeds its record limit.');
  const previous = map?.get(target.key),
    baseline = map
      ? previous === undefined
        ? null
        : (previous as string)
      : sequenceBaseline!.value;
  const vector = Y.decodeStateVector(Y.encodeStateVector(state.doc)),
    clock = Y.getState(state.doc.store, state.doc.clientID),
    id = previousIntent ?? state.doc.clientID + ':' + clock;
  const stamp = [
    1 + [...vector.values()].reduce((sum, n) => sum + n, 0),
    state.doc.clientID,
    clock,
  ] as const;
  const targetData = state.targets.has(key)
    ? null
    : encodeFieldRecord({
        ...target,
        baseline,
        ...(sequenceBaseline
          ? {
              position: sequenceBaseline.position,
              initiallyPresent: sequenceBaseline.initiallyPresent,
            }
          : {}),
      });
  if (targetData !== null) decodeFieldTarget(key, targetData);
  const intentData = encodeFieldRecord({ target: key, value, stamp });
  if (!previousIntent && state.intents.has(id))
    throw new ReviewFieldError('Review field intent identity was reused.');
  if (!previousIntent && reservation) reservation.remaining--;
  if (!state.targets.has('protocol')) set.call(state.targets, 'protocol', '1');
  if (targetData !== null) set.call(state.targets, key, targetData);
  state.origins.add(state.doc._transaction?.origin);
  set.call(state.intents, id, intentData);
  captured.set(key, id);
  if (map) {
    const keys = state.known.get(map) ?? new Map<string, string>();
    keys.set(target.key, key);
    state.known.set(map, keys);
  }
}
function install() {
  if (installed) return;
  installed = true;
  set = Y.Map.prototype.set;
  remove = Y.Map.prototype.delete;
  removeItems = Y.Array.prototype.delete;
  insertItems = Y.Array.prototype.insert;
  Y.Map.prototype.set = function (key, value) {
    const state = this.doc ? attached.get(this.doc) : undefined;
    if (state && typeof value === 'string') {
      const target = eligible(state, this, key, value);
      if (target) {
        if (!state.doc._transaction) {
          let result: typeof this = this;
          state.doc.transact(() => {
            result = this.set(key, value);
          });
          return result;
        }
        record(state, this, target, value);
      }
    } else if (
      state?.known.get(this)?.has(key) &&
      !state.normalizing &&
      state.doc._transaction?.local !== false
    ) {
      throw new ReviewFieldError('An owned review field cannot change its shared value type.');
    }
    return set.call(this, key, value);
  };
  Y.Map.prototype.delete = function (key) {
    const state = this.doc ? attached.get(this.doc) : undefined;
    if (state) {
      const target = eligible(state, this, key, null);
      if (target) {
        if (!state.doc._transaction) {
          state.doc.transact(() => {
            this.delete(key);
          });
          return;
        }
        record(state, this, target, null);
      }
    }
    return remove.call(this, key);
  };
  Y.Array.prototype.delete = function (index, length = 1) {
    const state = this.doc ? attached.get(this.doc) : undefined,
      transaction = this.doc?._transaction;
    if (
      state?.enabled &&
      !state.normalizing &&
      transaction?.local !== false &&
      !(transaction?.origin instanceof Y.UndoManager) &&
      (!state.captureOrigin || state.captureOrigin.value === transaction?.origin)
    ) {
      healthy(state);
      if (
        !Number.isSafeInteger(index) ||
        !Number.isSafeInteger(length) ||
        index < 0 ||
        length < 0 ||
        index + length > this.length
      )
        throw new ReviewFieldError('Invalid bounded review sequence removal.');
      if (!transaction) {
        state.doc.transact(() => this.delete(index, length));
        return;
      }
      const root = rootName(state, this);
      if (
        root !== null &&
        !excluded.has(root) &&
        (!this._item ||
          this._item.id.clock < (transaction.beforeState.get(this._item.id.client) ?? 0))
      )
        for (const entry of removedStringItems(this, index, length)) {
          const original = originalReviewSequenceTarget(
              state.doc,
              this,
              entry.key,
              records(state).targets.values()
            ),
            key = entry.key;
          record(
            state,
            null,
            original ??
              (this._item
                ? {
                    sequence: true,
                    client: this._item.id.client,
                    clock: this._item.id.clock,
                    key,
                  }
                : { sequence: true, root, key }),
            'removed',
            { value: entry.value, position: entry.position, initiallyPresent: true }
          );
        }
    }
    return removeItems.call(this, index, length);
  };
  Y.Array.prototype.insert = function (index, content) {
    const state = this.doc ? attached.get(this.doc) : undefined,
      transaction = this.doc?._transaction;
    if (
      !state?.enabled ||
      state.normalizing ||
      transaction?.local === false ||
      transaction?.origin instanceof Y.UndoManager ||
      (state.captureOrigin && state.captureOrigin.value !== transaction?.origin)
    )
      return insertItems.call(this, index, content);
    healthy(state);
    if (!transaction) {
      state.doc.transact(() => this.insert(index, content));
      return;
    }
    const root = rootName(state, this);
    if (
      root === null ||
      excluded.has(root) ||
      (this._item &&
        this._item.id.clock >= (transaction.beforeState.get(this._item.id.client) ?? 0))
    )
      return insertItems.call(this, index, content);
    const count = content.filter((value) => typeof value === 'string').length,
      reservation = state.reservations.at(-1);
    if (
      state.targets.size - Number(state.targets.has('protocol')) + count > state.limits.targets ||
      state.intents._map.size + count > state.limits.intents ||
      (reservation && count > reservation.remaining)
    )
      throw new ReviewFieldError('Review field ownership exceeds its complete insertion budget.');
    insertItems.call(this, index, content);
    for (const entry of removedStringItems(this, index, content.length))
      record(
        state,
        null,
        this._item
          ? {
              sequence: true,
              client: this._item.id.client,
              clock: this._item.id.clock,
              key: entry.key,
            }
          : { sequence: true, root, key: entry.key },
        'present',
        {
          value: entry.value,
          position: entry.position,
          initiallyPresent: false,
        }
      );
  };
}
function mapFor(state: State, target: ReviewFieldTarget): Y.Map<unknown> | null {
  if (target.root !== undefined) {
    if (!state.doc.share.has(target.root)) return null;
    return state.doc.getMap(target.root);
  }
  return flagMap(state.doc, {
    client: target.client!,
    clock: target.clock!,
    key: target.key,
    baseline: 'missing',
  });
}
function records(state: State) {
  if (
    state.targets.size > state.limits.targets + 1 ||
    state.intents._map.size > state.limits.intents
  )
    throw new ReviewFieldError('Review field ownership exceeds its record limit.');
  const version = state.targets.get('protocol');
  if (version === undefined && !state.targets.size && !state.intents.size)
    return {
      targets: new Map<string, ReviewFieldTarget>(),
      winners: new Map<string, ReviewFieldIntent>(),
    };
  if (version !== '1') throw new ReviewFieldError('Unsupported owned review field protocol.');
  const targets = new Map<string, ReviewFieldTarget>();
  for (const [key, raw] of state.targets)
    if (key !== 'protocol') {
      const target = decodeFieldTarget(key, raw);
      if (target.root !== undefined && excluded.has(target.root))
        throw new ReviewFieldError('Review fields cannot target ownership metadata.');
      const resolved = target.sequence ? reviewArray(state.doc, target) : mapFor(state, target);
      const targetRoot = resolved ? rootName(state, resolved) : null;
      if (targetRoot !== null && excluded.has(targetRoot))
        throw new ReviewFieldError('Review fields cannot target ownership metadata.');
      targets.set(key, target);
    }
  const winners = new Map<string, ReviewFieldIntent>();
  for (const [key, raw] of state.intents) {
    const intent = decodeFieldIntent(key, raw);
    if (!targets.has(intent.target))
      throw new ReviewFieldError('Review field intent has no target.');
    if (targets.get(intent.target)?.sequence && !['removed', 'present'].includes(intent.value!))
      throw new ReviewFieldError('Invalid review sequence intent.');
    const old = winners.get(intent.target);
    if (!old || newerFieldIntent(intent, old)) winners.set(intent.target, intent);
  }
  return { targets, winners };
}
function normalize(state: State) {
  const { targets, winners } = records(state),
    edits: (() => void)[] = [],
    arrays = new Map<
      Y.Array<unknown>,
      { target: ReviewFieldTarget; present: boolean; restoreOrder: boolean }[]
    >();
  for (const [key, target] of targets) {
    if (target.sequence) {
      const array = reviewArray(state.doc, target);
      if (array) {
        const keys = arrays.get(array) ?? [];
        const winner = winners.get(key);
        const present = winner ? winner.value === 'present' : target.initiallyPresent!;
        keys.push({
          target,
          present,
          restoreOrder: state.sequencePresence.get(key) === false && present,
        });
        state.sequencePresence.set(key, present);
        arrays.set(array, keys);
      }
      continue;
    }
    const map = mapFor(state, target);
    if (!map) continue;
    const root = rootName(state, map);
    if (root !== null && excluded.has(root))
      throw new ReviewFieldError('Review fields cannot target ownership metadata.');
    const keys = state.known.get(map) ?? new Map<string, string>();
    keys.set(target.key, key);
    state.known.set(map, keys);
    const winner = winners.get(key),
      desired = winner ? winner.value : target.baseline,
      current = map.get(target.key);
    if (current !== undefined && typeof current !== 'string')
      throw new ReviewFieldError('Review field changed its shared value type.');
    if (desired === null ? map.has(target.key) : current !== desired)
      edits.push(() =>
        desired === null ? remove.call(map, target.key) : set.call(map, target.key, desired)
      );
  }
  if (edits.length || arrays.size)
    state.doc.transact(() => {
      for (const edit of edits) edit();
      for (const [array, keys] of arrays)
        normalizeReviewArray(state.doc, array, keys, removeItems, insertItems);
    }, NORMALIZE);
}
function limits(options: OwnedReviewFieldOptions) {
  const result = {
    targets: options.limits?.targets ?? MAX_REVIEW_FIELD_TARGETS,
    intents: options.limits?.intents ?? MAX_REVIEW_FIELD_INTENTS,
  };
  if (
    !Number.isSafeInteger(result.targets) ||
    result.targets < 0 ||
    result.targets > MAX_REVIEW_FIELD_TARGETS ||
    !Number.isSafeInteger(result.intents) ||
    result.intents < 0 ||
    result.intents > MAX_REVIEW_FIELD_INTENTS
  )
    throw new ReviewFieldError('Invalid bounded review field limits.');
  return result;
}
export function attachOwnedReviewFields(doc: Y.Doc, options: OwnedReviewFieldOptions = {}) {
  if (attached.has(doc)) throw new ReviewFieldError('Review fields are already attached.');
  const state: State = {
    doc,
    targets: doc.getMap(REVIEW_FIELD_TARGETS),
    intents: doc.getMap(REVIEW_FIELD_INTENTS),
    origins: new Set(),
    known: new WeakMap(),
    transactions: new WeakMap(),
    sequencePresence: new Map(),
    reservations: [],
    limits: limits(options),
    enabled: false,
    capturing: false,
    captureOrigin: null,
    normalizing: false,
    deferred: false,
    disposed: false,
    error: null,
  };
  const fail = (cause: unknown) => {
    state.error = cause instanceof Error ? cause : new ReviewFieldError(String(cause));
    options.onError?.(state.error);
    throw state.error;
  };
  if (!doc.store.pendingStructs && !doc.store.pendingDs) records(state);
  install();
  attached.set(doc, state);
  const beforeTransaction = (transaction: Y.Transaction) => {
    if (state.capturing && state.captureOrigin === null)
      state.captureOrigin = { value: transaction.origin };
    const manager = transaction.origin;
    if (
      manager instanceof Y.UndoManager &&
      [...state.origins].some((origin) => manager.trackedOrigins.has(origin))
    )
      manager.addToScope(state.intents);
  };
  const beforeObservers = (transaction: Y.Transaction) => {
    if (state.normalizing || transaction.origin === NORMALIZE) return;
    try {
      healthy(state);
      if (doc.store.pendingStructs || doc.store.pendingDs) {
        state.deferred = true;
        return;
      }
      const changed = new Set<unknown>(transaction.changed.keys());
      if (
        !state.deferred &&
        !changed.has(state.targets) &&
        !changed.has(state.intents) &&
        !(transaction.origin instanceof Y.UndoManager) &&
        ![...transaction.changed].some(
          ([type, keys]) =>
            type instanceof Y.Map &&
            [...keys].some((key) => key !== null && state.known.get(type)?.has(key))
        )
      )
        return;
      state.deferred = false;
      state.normalizing = true;
      recordRestoredLineage(doc, transaction, NORMALIZE);
      normalize(state);
    } catch (error) {
      fail(error);
    } finally {
      state.normalizing = false;
    }
  };
  doc.on('beforeTransaction', beforeTransaction);
  doc.on('beforeObserverCalls', beforeObservers);
  return Object.freeze({
    enable() {
      healthy(state);
      state.enabled = true;
    },
    assertHealthy() {
      healthy(state);
    },
    capture<T>(run: () => T, fieldWrites: number): T {
      healthy(state);
      if (doc.store.pendingStructs || doc.store.pendingDs)
        throw new ReviewFieldError('Review fields must wait for complete room updates.');
      if (!Number.isSafeInteger(fieldWrites) || fieldWrites < 0)
        throw new ReviewFieldError('Review ownership needs a complete bounded field count.');
      records(state);
      const reserved = state.reservations.reduce((sum, item) => sum + item.remaining, 0),
        targets = state.targets.size - Number(state.targets.has('protocol'));
      if (
        targets + reserved + fieldWrites > state.limits.targets ||
        state.intents._map.size + reserved + fieldWrites > state.limits.intents
      )
        throw new ReviewFieldError('Review field ownership exceeds its record limit.');
      state.reservations.push({ remaining: fieldWrites, captured: new Map() });
      const previous = state.capturing,
        origin = state.captureOrigin;
      state.capturing = true;
      state.captureOrigin = null;
      try {
        return run();
      } finally {
        state.reservations.pop();
        state.capturing = previous;
        state.captureOrigin = origin;
      }
    },
    dispose() {
      if (state.disposed) return;
      state.disposed = true;
      doc.off('beforeTransaction', beforeTransaction);
      doc.off('beforeObserverCalls', beforeObservers);
      state.origins.clear();
      state.sequencePresence.clear();
      attached.delete(doc);
    },
  });
}
