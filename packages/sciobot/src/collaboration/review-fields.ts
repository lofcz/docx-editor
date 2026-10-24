import { PROPERTY_GROUP_RECEIPTS } from './property-group-codec';
import { nativeBirthFieldCreation } from './review-birth-field-retention';
import {
  retainedNumberingBirths,
  nativeNumberingChildArray,
} from './review-numbering-birth-retention';
// Application-owned review string fields over pinned MIT Yjs.
// Pinned retained-birth interoperability stays in review-sequence-retention.ts.
import * as Y from 'yjs';
import { OWNED_TEXT_OPERATIONS, OWNED_TEXT_LINEAGE } from './ownership-codec';
import {
  REVIEW_FLAG_TARGETS,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_RETRACTIONS,
} from './review-flag-codec';
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
  REVIEW_FIELD_RETRACTIONS,
  MAX_REVIEW_FIELD_TARGETS,
  MAX_REVIEW_FIELD_INTENTS,
  MAX_REVIEW_FIELD_BYTES,
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
  /** Experimental protocol 2 requires explicit attachment; existing hosts keep protocol 1. */
  readonly protocol?: '1' | '2';
  /** Hosts may lower the fixed record limits. Raising them is refused. */
  readonly limits?: { readonly targets?: number; readonly intents?: number };
}
interface Reservation {
  remaining: number;
  extraRemaining: number;
  readonly captured: Map<string, string>;
  readonly retracted?: ReadonlySet<string>;
}
/** Exact local publication evidence. Persisting or binding this evidence needs a separate protocol. */
export interface ReviewFieldReceipt {
  readonly id: string;
  readonly record: string;
}
interface State {
  readonly protocol: '1' | '2';
  readonly doc: Y.Doc;
  readonly targets: Y.Map<string>;
  readonly intents: Y.Map<string>;
  readonly retractions: Y.Map<string> | null;
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
function retainedIntentCount(state: State): number {
  return state.intents._map.size + (state.retractions?._map.size ?? 0);
}
function decodeFieldRetraction(id: string, raw: unknown): ReviewFieldReceipt {
  if (typeof raw !== 'string' || new TextEncoder().encode(raw).length > MAX_REVIEW_FIELD_BYTES)
    throw new ReviewFieldError('Invalid bounded field retraction.');
  const item = JSON.parse(raw);
  if (
    !item ||
    typeof item !== 'object' ||
    Array.isArray(item) ||
    Object.keys(item).length !== 2 ||
    typeof item.id !== 'string' ||
    typeof item.record !== 'string'
  )
    throw new ReviewFieldError('Invalid field retraction receipt.');
  decodeFieldIntent(id, item.record);
  decodeFieldIntent(item.id, item.record);
  return item;
}
const NORMALIZE = Symbol('sciobot-review-field-normalization');
const attached = new WeakMap<Y.Doc, State>();
const excluded = new Set([
  PROPERTY_GROUP_RECEIPTS,
  REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
  REVIEW_FIELD_RETRACTIONS,
  REVIEW_FLAG_TARGETS,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_RETRACTIONS,
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
  // The native birth delete filter owns initial geometry and Word identities. Layering
  // an absent scalar baseline over them would erase retained peer-bearing cell structure.
  if (
    transaction &&
    nativeBirthFieldCreation(state.doc, transaction, target, state.protocol === '2')
  )
    return null;
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
  // A rejection withdraws its own intent. Recording the native snapshot restoration
  // as another write would overwrite the peer whose intent must now win.
  if (state.reservations.some((scope) => scope.retracted?.has(key))) return;
  if (!previousIntent && reservation && !reservation.remaining)
    throw new ReviewFieldError('The review command exceeded its planned field count.');
  if (
    (!state.targets.has(key) &&
      state.targets.size - Number(state.targets.has('protocol')) >= state.limits.targets) ||
    (!previousIntent && retainedIntentCount(state) >= state.limits.intents)
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
  if (!state.targets.has('protocol')) set.call(state.targets, 'protocol', state.protocol);
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
        this._item.id.clock >= (transaction.beforeState.get(this._item.id.client) ?? 0) &&
        !(state.protocol === '2' && nativeNumberingChildArray(state.doc, this)))
    )
      return insertItems.call(this, index, content);
    const count = content.filter((value) => typeof value === 'string').length,
      reservation = state.reservations.at(-1);
    if (
      state.targets.size - Number(state.targets.has('protocol')) + count > state.limits.targets ||
      retainedIntentCount(state) + count > state.limits.intents ||
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
    retainedIntentCount(state) > state.limits.intents
  )
    throw new ReviewFieldError('Review field ownership exceeds its record limit.');
  const version = state.targets.get('protocol');
  if (
    version === undefined &&
    !state.targets.size &&
    !state.intents.size &&
    !state.retractions?.size
  )
    return {
      targets: new Map<string, ReviewFieldTarget>(),
      winners: new Map<string, ReviewFieldIntent>(),
    };
  if (version !== state.protocol)
    throw new ReviewFieldError('Unsupported owned review field protocol.');
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
  const withdrawn = new Set<string>();
  for (const [id, raw] of state.retractions ?? []) {
    const receipt = decodeFieldRetraction(id, raw);
    if (state.intents.has(receipt.id) && state.intents.get(receipt.id) !== receipt.record)
      throw new ReviewFieldError('Field retraction does not match its original intent.');
    const intent = decodeFieldIntent(receipt.id, receipt.record);
    if (!targets.has(intent.target))
      throw new ReviewFieldError('Field retraction must target a known intent.');
    withdrawn.add(receipt.id);
  }
  const winners = new Map<string, ReviewFieldIntent>();
  for (const [key, raw] of state.intents) {
    const intent = decodeFieldIntent(key, raw);
    if (!targets.has(intent.target))
      throw new ReviewFieldError('Review field intent has no target.');
    if (targets.get(intent.target)?.sequence && !['removed', 'present'].includes(intent.value!))
      throw new ReviewFieldError('Invalid review sequence intent.');
    if (withdrawn.has(key)) continue;
    const old = winners.get(intent.target);
    if (!old || newerFieldIntent(intent, old)) winners.set(intent.target, intent);
  }
  return { targets, winners };
}
function normalize(state: State) {
  const { targets, winners } = records(state),
    retained = state.protocol === '2' ? retainedNumberingBirths(state.doc, targets, winners) : null,
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
        const present = winner
          ? winner.value === 'present'
          : target.initiallyPresent! || retained?.has(key) === true;
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
        normalizeReviewArray(
          state.doc,
          array,
          keys,
          removeItems,
          insertItems,
          state.protocol === '2'
        );
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
  if (options.protocol !== undefined && !['1', '2'].includes(options.protocol))
    throw new ReviewFieldError('Unsupported owned review field protocol.');
  const protocol = options.protocol ?? '1';
  // Cold structural receipts need restored item ancestry before their first normalization.
  if (protocol === '2') doc.getMap<string>(OWNED_TEXT_LINEAGE);
  const state: State = {
    protocol,
    doc,
    targets: doc.getMap(REVIEW_FIELD_TARGETS),
    intents: doc.getMap(REVIEW_FIELD_INTENTS),
    retractions: protocol === '2' ? doc.getMap(REVIEW_FIELD_RETRACTIONS) : null,
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
    ) {
      manager.addToScope(state.intents);
      if (state.retractions) manager.addToScope(state.retractions);
    }
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
      // Protocol 2 waits for ownership metadata before correcting remote physical writes.
      // Otherwise disconnected replicas can exchange opposing corrections before their intents arrive.
      if (
        !state.deferred &&
        !changed.has(state.targets) &&
        !changed.has(state.intents) &&
        !changed.has(state.retractions) &&
        !(state.protocol === '2' && changed.has(doc.share.get(OWNED_TEXT_LINEAGE))) &&
        !(transaction.origin instanceof Y.UndoManager) &&
        !(
          (state.protocol === '1' || transaction.local) &&
          [...transaction.changed].some(
            ([type, keys]) =>
              type instanceof Y.Map &&
              [...keys].some((key) => key !== null && state.known.get(type)?.has(key))
          )
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
  const capture = <T>(
    run: () => T,
    fieldWrites: number,
    receipt?: (items: readonly ReviewFieldReceipt[]) => void,
    retracted?: ReadonlySet<string>,
    extraIntents = 0
  ): T => {
    healthy(state);
    if (doc.store.pendingStructs || doc.store.pendingDs)
      throw new ReviewFieldError('Review fields must wait for complete room updates.');
    if (!Number.isSafeInteger(fieldWrites) || fieldWrites < 0)
      throw new ReviewFieldError('Review ownership needs a complete bounded field count.');
    records(state);
    const reserved = state.reservations.reduce((sum, item) => sum + item.remaining, 0),
      reservedIntents = state.reservations.reduce(
        (sum, item) => sum + item.remaining + item.extraRemaining,
        0
      ),
      targets = state.targets.size - Number(state.targets.has('protocol'));
    if (
      targets + reserved + fieldWrites > state.limits.targets ||
      retainedIntentCount(state) + reservedIntents + fieldWrites + extraIntents >
        state.limits.intents
    )
      throw new ReviewFieldError('Review field ownership exceeds its record limit.');
    const reservation: Reservation = {
      remaining: fieldWrites,
      extraRemaining: extraIntents,
      captured: new Map(),
      retracted,
    };
    state.reservations.push(reservation);
    const previous = state.capturing,
      origin = state.captureOrigin;
    state.capturing = true;
    state.captureOrigin = null;
    try {
      const result = run();
      if (receipt)
        receipt(
          Object.freeze(
            [...reservation.captured.values()].map((id) =>
              Object.freeze({ id, record: state.intents.get(id)! })
            )
          )
        );
      return result;
    } finally {
      state.reservations.pop();
      state.capturing = previous;
      state.captureOrigin = origin;
    }
  };
  const prepareRetraction = (items: readonly ReviewFieldReceipt[], fieldWrites = 0) => {
    healthy(state);
    if (!state.retractions)
      throw new ReviewFieldError('Field retraction requires explicit protocol 2 attachment.');
    if (!Array.isArray(items) || items.length > state.limits.intents)
      throw new ReviewFieldError('Invalid bounded field retraction receipt.');
    if (!Number.isSafeInteger(fieldWrites) || fieldWrites < 0)
      throw new ReviewFieldError('Review ownership needs a complete bounded field count.');
    records(state);
    const withdrawn = new Set(
      [...state.retractions].map(([id, raw]) => decodeFieldRetraction(id, raw).id)
    );
    const ids = new Set<string>(),
      targets = new Set<string>(),
      encoded: string[] = [];
    for (const item of items) {
      if (
        !item ||
        typeof item.id !== 'string' ||
        typeof item.record !== 'string' ||
        Object.keys(item).length !== 2 ||
        ids.has(item.id) ||
        withdrawn.has(item.id) ||
        state.intents.get(item.id) !== item.record
      )
        throw new ReviewFieldError('Field retraction receipt is stale or invalid.');
      const intent = decodeFieldIntent(item.id, item.record);
      decodeFieldTarget(intent.target, state.targets.get(intent.target));
      ids.add(item.id);
      targets.add(intent.target);
      encoded.push(encodeFieldRecord(item));
    }

    return { targets, encoded };
  };
  return Object.freeze({
    enable() {
      healthy(state);
      state.enabled = true;
    },
    assertHealthy() {
      healthy(state);
    },
    capture<T>(
      run: () => T,
      fieldWrites: number,
      receipt?: (items: readonly ReviewFieldReceipt[]) => void
    ): T {
      return capture(run, fieldWrites, receipt);
    },
    currentReceipt(): readonly ReviewFieldReceipt[] {
      const reservation = state.reservations.at(-1);
      return Object.freeze(
        [...(reservation?.captured.values() ?? [])].map((id) =>
          Object.freeze({ id, record: state.intents.get(id)! })
        )
      );
    },
    assertRetraction(items: readonly ReviewFieldReceipt[], fieldWrites = 0): void {
      const { targets } = prepareRetraction(items, fieldWrites);
      capture(() => {}, fieldWrites, undefined, targets, items.length);
    },
    /** Internal ownership primitive. The caller must own the active native transaction. */
    retractInTransaction<T>(
      items: readonly ReviewFieldReceipt[],
      run: () => T,
      fieldWrites = 0
    ): T {
      const { targets, encoded } = prepareRetraction(items, fieldWrites);
      if (!doc._transaction?.local || doc._transaction.origin instanceof Y.UndoManager)
        throw new ReviewFieldError('Field retraction requires an active local native transaction.');
      return capture(
        () => {
          state.origins.add(doc._transaction!.origin);
          for (const raw of encoded) {
            const id = `${doc.clientID}:${Y.getState(doc.store, doc.clientID)}`;
            state.reservations.at(-1)!.extraRemaining--;
            set.call(state.retractions!, id, raw);
          }
          return run();
        },
        fieldWrites,
        undefined,
        targets,
        items.length
      );
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
