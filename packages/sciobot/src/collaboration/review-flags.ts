import * as Y from 'yjs';
import type { OwnedTextOptions } from './owned-text';
import {
  decodeFlagTarget,
  flagAddress,
  MAX_REVIEW_FLAGS,
  MAX_REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_TARGETS,
  ReviewFlagError,
  validFlagKey,
  validateFlagIntent,
  type ReviewFlagTarget,
} from './review-flag-codec';
import { flagMap } from './review-flag-items';

export interface OwnedReviewFlagOptions extends OwnedTextOptions {
  readonly protocol?: '1' | '2';
}

interface State {
  readonly protocol: '1' | '2';
  readonly doc: Y.Doc;
  readonly targets: Y.Map<unknown>;
  readonly intents: Y.Map<unknown>;
  readonly origins: Set<unknown>;
  readonly known: WeakMap<Y.Map<unknown>, Set<string>>;
  error: Error | null;
  capturing: boolean;
  captureOrigin: { readonly value: unknown } | null;
  readonly reservations: { remaining: number; readonly captured: Set<string> }[];
  normalizing: boolean;
  disposed: boolean;
  deferred: boolean;
}
export interface OwnedReviewFlags {
  assertHealthy(): void;
  /** The trusted canonical preview supplies an upper bound for distinct removal flags. */
  capture<T>(run: () => T, flagWrites: number): T;
  dispose(): void;
}
const attached = new WeakMap<Y.Doc, State>();
const NORMALIZE = Symbol('sciobot-review-flag-normalization');
const set = Y.Map.prototype.set;
let patched = false;

function healthy(state: State) {
  if (state.disposed || state.doc.isDestroyed)
    throw new ReviewFlagError('The review flag scope is no longer active.');
  if (state.error) throw state.error;
}
function retainedIntents(state: State): number {
  // Pinned MIT Yjs keeps a map key's item after undo. Redo reuses that same identity.
  // Visible size alone would let another actor spend the original actor's redo capacity.
  return state.intents._map.size;
}
function captureFlag(state: State, map: Y.Map<unknown>, key: string) {
  healthy(state);
  if (!state.doc._transaction)
    throw new ReviewFlagError('Review flags must share the native transaction.');
  if (!validFlagKey(key) || !map._item)
    throw new ReviewFlagError('A review flag needs a bounded shared map address.');
  const existing = map.get(key);
  if (existing !== undefined && existing !== false && existing !== true)
    throw new ReviewFlagError('A review flag cannot replace a non-boolean field.');
  const target: ReviewFlagTarget = {
      client: map._item.id.client,
      clock: map._item.id.clock,
      key,
      baseline: existing === false ? 'false' : 'missing',
    },
    address = flagAddress(target);
  const reservation = state.reservations.at(-1);
  if (!reservation) throw new ReviewFlagError('Review flags need a complete capacity reservation.');
  if (reservation.captured.has(address)) return;
  if (existing === true && !state.targets.has(address)) return;
  if (!reservation.remaining)
    throw new ReviewFlagError('The review command exceeded its planned removal flag count.');
  if (
    (!state.targets.has(address) &&
      state.targets.size - Number(state.targets.has('protocol')) >= MAX_REVIEW_FLAGS) ||
    retainedIntents(state) >= MAX_REVIEW_FLAG_INTENTS
  )
    throw new ReviewFlagError('Review flag ownership exceeds its record limit.');
  reservation.remaining--;
  reservation.captured.add(address);
  state.doc.transact((transaction) => {
    state.origins.add(transaction.origin);
    if (!state.targets.has('protocol')) set.call(state.targets, 'protocol', state.protocol);
    if (!state.targets.has(address)) {
      // A pre-existing true flag has no pending removal to own.
      if (existing === true) return;
      set.call(state.targets, address, JSON.stringify(target));
    }
    const clock = Y.getState(state.doc.store, state.doc.clientID),
      id = `${state.doc.clientID}:${clock}`;
    if (state.intents.has(id)) throw new ReviewFlagError('Review flag intent identity was reused.');
    set.call(state.intents, id, address);
  });
}
function install() {
  if (patched) return;
  patched = true;
  Y.Map.prototype.set = function (key, value) {
    const state = this.doc ? attached.get(this.doc) : undefined;
    if (
      state?.capturing &&
      !state.normalizing &&
      value === true &&
      this._item &&
      (!state.captureOrigin || state.captureOrigin.value === state.doc._transaction?.origin)
    )
      captureFlag(state, this, key);
    return set.call(this, key, value);
  };
}
function normalize(state: State) {
  if (state.targets.size > MAX_REVIEW_FLAGS + 1 || retainedIntents(state) > MAX_REVIEW_FLAG_INTENTS)
    throw new ReviewFlagError('Review flag ownership exceeds its record limit.');
  const version = state.targets.get('protocol');
  if (version === undefined && !state.targets.size && !state.intents.size) return;
  if (version !== state.protocol) throw new ReviewFlagError('Unsupported review flag protocol.');
  if (state.protocol === '2') {
    let item = state.targets._map.get('protocol');
    let count = 0;
    while (item) {
      if (++count > 4096) throw new ReviewFlagError('Review protocol history exceeds its limit.');
      if (
        item.parentSub === 'protocol' &&
        item.content.getContent().some((value) => value === '1')
      ) {
        // Keep a mixed-room refusal in snapshots after Yjs collects the losing header.
        set.call(state.targets, 'protocol', 'mixed');
        throw new ReviewFlagError('Legacy review history requires export and reseed.');
      }
      item = item.left ?? undefined;
    }
  }
  const targets = new Map<string, ReviewFlagTarget>(),
    active = new Set<string>();
  for (const [address, data] of state.targets)
    if (address !== 'protocol') targets.set(address, decodeFlagTarget(address, data));
  for (const [id, value] of state.intents) {
    validateFlagIntent(id, value);
    if (!targets.has(value)) throw new ReviewFlagError('Review flag intent has no target.');
    active.add(value);
  }
  const edits: (() => void)[] = [];
  for (const [address, target] of targets) {
    const map = flagMap(state.doc, target);
    if (!map) continue;
    const keys = state.known.get(map) ?? new Set<string>();
    keys.add(target.key);
    state.known.set(map, keys);
    const value = map.get(target.key);
    if (value !== undefined && value !== true && value !== false)
      throw new ReviewFlagError('A review flag field contains a non-boolean value.');
    const desired = active.has(address) ? true : target.baseline === 'false' ? false : undefined;
    if (value !== desired || (desired === undefined && map.has(target.key)))
      edits.push(() =>
        desired === undefined ? map.delete(target.key) : set.call(map, target.key, desired)
      );
  }
  // No empty transaction: observers must not schedule each other forever.
  if (edits.length)
    state.doc.transact(() => {
      for (const edit of edits) edit();
    }, NORMALIZE);
}

/** Own true-valued map flags written by a bounded review command, using only pinned MIT Yjs. */
export function attachOwnedReviewFlags(
  doc: Y.Doc,
  options: OwnedReviewFlagOptions
): OwnedReviewFlags {
  const protocol = options.protocol ?? '1';
  if (protocol !== '1' && protocol !== '2')
    throw new ReviewFlagError('Unsupported review flag protocol.');
  if (attached.has(doc)) throw new ReviewFlagError('Review flags are already attached.');
  const state: State = {
    protocol,
    doc,
    targets: doc.getMap(REVIEW_FLAG_TARGETS),
    intents: doc.getMap(REVIEW_FLAG_INTENTS),
    origins: new Set(),
    known: new WeakMap(),
    error: null,
    capturing: false,
    captureOrigin: null,
    reservations: [],
    normalizing: false,
    disposed: false,
    deferred: false,
  };
  const storedProtocol = state.targets.get('protocol');
  if (storedProtocol !== undefined && storedProtocol !== protocol)
    throw new ReviewFlagError('Unsupported review flag protocol.');
  if (storedProtocol === undefined && (state.targets.size || state.intents.size))
    throw new ReviewFlagError('Review flag records need their protocol.');
  install();
  attached.set(doc, state);
  const fail = (cause: unknown): never => {
    state.error = cause instanceof Error ? cause : new ReviewFlagError(String(cause));
    options.onError?.(state.error);
    throw state.error;
  };
  const beforeTransaction = (transaction: Y.Transaction) => {
    // A synchronous relay or materializer transaction is not part of this local command.
    if (state.capturing && state.captureOrigin === null)
      state.captureOrigin = { value: transaction.origin };
    const manager = transaction.origin;
    if (
      manager instanceof Y.UndoManager &&
      [...state.origins].some((origin) => manager.trackedOrigins.has(origin))
    ) {
      // The original frame already contains every inserted ID, including these intent rows.
      // Add their root before the manager pops that frame. No private provider field is read.
      manager.addToScope(state.intents);
    }
  };
  const beforeObservers = (transaction: Y.Transaction) => {
    if (transaction.origin === NORMALIZE || state.normalizing) return;
    try {
      healthy(state);
      // Wait for complete deliveries, including an intent arriving after its normalization.
      if (doc.store.pendingStructs || doc.store.pendingDs) {
        state.deferred = true;
        return;
      }
      const changed = new Set<unknown>(transaction.changed.keys());
      const relevant =
        state.deferred ||
        changed.has(state.targets) ||
        changed.has(state.intents) ||
        transaction.origin instanceof Y.UndoManager ||
        [...transaction.changed].some(
          ([type, keys]) =>
            type instanceof Y.Map &&
            [...keys].some((key) => key !== null && state.known.get(type)?.has(key))
        );
      if (!relevant) return;
      state.deferred = false;
      state.normalizing = true;
      normalize(state);
    } catch (error) {
      fail(error);
    } finally {
      state.normalizing = false;
    }
  };
  doc.on('beforeTransaction', beforeTransaction);
  doc.on('beforeObserverCalls', beforeObservers);
  if (protocol === '2' && storedProtocol === undefined)
    set.call(state.targets, 'protocol', protocol);
  return Object.freeze({
    assertHealthy: () => healthy(state),
    capture<T>(run: () => T, flagWrites: number): T {
      healthy(state);
      if (doc.store.pendingStructs || doc.store.pendingDs)
        throw new ReviewFlagError('Review decisions must wait for complete room updates.');
      if (!Number.isSafeInteger(flagWrites) || flagWrites < 0)
        throw new ReviewFlagError('Review ownership needs a complete bounded removal flag count.');
      const reserved = state.reservations.reduce((sum, item) => sum + item.remaining, 0),
        targets = state.targets.size - Number(state.targets.has('protocol'));
      // Reserve the whole command before its callback can mutate canonical or shared state.
      if (
        targets + reserved + flagWrites > MAX_REVIEW_FLAGS ||
        retainedIntents(state) + reserved + flagWrites > MAX_REVIEW_FLAG_INTENTS
      )
        throw new ReviewFlagError('Review flag ownership exceeds its record limit.');
      state.reservations.push({ remaining: flagWrites, captured: new Set() });
      const previous = state.capturing;
      const previousOrigin = state.captureOrigin;
      state.capturing = true;
      state.captureOrigin = null;
      try {
        return run();
      } finally {
        state.reservations.pop();
        state.capturing = previous;
        state.captureOrigin = previousOrigin;
      }
    },
    dispose() {
      if (state.disposed) return;
      state.disposed = true;
      doc.off('beforeTransaction', beforeTransaction);
      doc.off('beforeObserverCalls', beforeObservers);
      state.origins.clear();
      attached.delete(doc);
    },
  });
}
