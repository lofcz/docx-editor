import * as Y from 'yjs';
import type { OwnedTextOptions } from './owned-text';
import {
  decodeFlagTarget,
  flagAddress,
  MAX_REVIEW_FLAGS,
  MAX_REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_RETRACTIONS,
  decodeFlagRetraction,
  type ReviewFlagReceipt,
  REVIEW_FLAG_TARGETS,
  ReviewFlagError,
  validFlagKey,
  validateFlagIntent,
  type ReviewFlagTarget,
} from './review-flag-codec';
import { flagMap } from './review-flag-items';

export interface OwnedReviewFlagOptions extends OwnedTextOptions {
  readonly protocol?: '1' | '2' | '3';
  readonly bootstrap?: 'create' | 'join';
}

interface State {
  readonly protocol: '1' | '2' | '3';
  readonly doc: Y.Doc;
  readonly targets: Y.Map<unknown>;
  readonly intents: Y.Map<unknown>;
  readonly retractions: Y.Map<string> | null;
  readonly origins: Set<unknown>;
  readonly known: WeakMap<Y.Map<unknown>, Set<string>>;
  error: Error | null;
  capturing: boolean;
  captureOrigin: { readonly value: unknown } | null;
  readonly reservations: {
    remaining: number;
    readonly captured: Map<string, string>;
    readonly retracted?: ReadonlySet<string>;
    extraRemaining: number;
  }[];
  normalizing: boolean;
  disposed: boolean;
  deferred: boolean;
}
export interface OwnedReviewFlags {
  assertHealthy(): void;
  /** The trusted canonical preview supplies an upper bound for distinct removal flags. */
  capture<T>(
    run: () => T,
    flagWrites: number,
    receipt?: (items: readonly ReviewFlagReceipt[]) => void
  ): T;
  assertRetraction(items: readonly ReviewFlagReceipt[], flagWrites?: number): void;
  retractInTransaction<T>(
    items: readonly ReviewFlagReceipt[],
    run: () => T,
    flagWrites?: number
  ): T;
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
  return state.intents._map.size + (state.retractions?._map.size ?? 0);
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
  if (state.reservations.some((scope) => scope.retracted?.has(address))) return;
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
    reservation.captured.set(address, id);
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
  if (state.protocol !== '1') {
    let item = state.targets._map.get('protocol');
    let count = 0;
    while (item) {
      if (++count > 4096) throw new ReviewFlagError('Review protocol history exceeds its limit.');
      if (
        item.parentSub === 'protocol' &&
        item.content
          .getContent()
          .some((value) => value === '1' || (state.protocol === '3' && value === '2'))
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
  const withdrawn = new Set<string>();
  for (const [id, raw] of state.retractions ?? []) {
    const receipt = decodeFlagRetraction(id, raw);
    if (
      !targets.has(receipt.record) ||
      (state.intents.has(receipt.id) && state.intents.get(receipt.id) !== receipt.record)
    )
      throw new ReviewFlagError('Flag retraction does not match its original intent.');
    withdrawn.add(receipt.id);
  }
  for (const [id, value] of state.intents) {
    validateFlagIntent(id, value);
    if (!targets.has(value)) throw new ReviewFlagError('Review flag intent has no target.');
    if (!withdrawn.has(id)) active.add(value);
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
  if (protocol !== '1' && protocol !== '2' && protocol !== '3')
    throw new ReviewFlagError('Unsupported review flag protocol.');
  if (attached.has(doc)) throw new ReviewFlagError('Review flags are already attached.');
  const state: State = {
    protocol,
    doc,
    targets: doc.getMap(REVIEW_FLAG_TARGETS),
    intents: doc.getMap(REVIEW_FLAG_INTENTS),
    retractions: protocol === '3' ? doc.getMap(REVIEW_FLAG_RETRACTIONS) : null,
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
      if (state.retractions) manager.addToScope(state.retractions);
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
        changed.has(state.retractions) ||
        transaction.origin instanceof Y.UndoManager ||
        ((protocol !== '3' || transaction.local) &&
          [...transaction.changed].some(
            ([type, keys]) =>
              type instanceof Y.Map &&
              [...keys].some((key) => key !== null && state.known.get(type)?.has(key))
          ));
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
  if (protocol !== '1' && storedProtocol === undefined && options.bootstrap !== 'join')
    set.call(state.targets, 'protocol', protocol);
  const capture = <T>(
    run: () => T,
    flagWrites: number,
    receipt?: (items: readonly ReviewFlagReceipt[]) => void,
    retracted?: ReadonlySet<string>,
    extraIntents = 0
  ): T => {
    healthy(state);
    if (doc.store.pendingStructs || doc.store.pendingDs)
      throw new ReviewFlagError('Review decisions must wait for complete room updates.');
    if (!Number.isSafeInteger(flagWrites) || flagWrites < 0)
      throw new ReviewFlagError('Review ownership needs a complete bounded removal flag count.');
    const reserved = state.reservations.reduce((sum, item) => sum + item.remaining, 0),
      reservedIntents = state.reservations.reduce(
        (sum, item) => sum + item.remaining + item.extraRemaining,
        0
      ),
      targets = state.targets.size - Number(state.targets.has('protocol'));
    // Reserve the whole command before its callback can mutate canonical or shared state.
    if (
      targets + reserved + flagWrites > MAX_REVIEW_FLAGS ||
      retainedIntents(state) + reservedIntents + flagWrites + extraIntents > MAX_REVIEW_FLAG_INTENTS
    )
      throw new ReviewFlagError('Review flag ownership exceeds its record limit.');
    const reservation = {
      remaining: flagWrites,
      extraRemaining: extraIntents,
      captured: new Map<string, string>(),
      retracted,
    };
    state.reservations.push(reservation);
    const previous = state.capturing;
    const previousOrigin = state.captureOrigin;
    state.capturing = true;
    state.captureOrigin = null;
    try {
      const result = run();
      if (receipt)
        receipt(
          Object.freeze(
            [...reservation.captured.values()].map((id) =>
              Object.freeze({ id, record: state.intents.get(id) as string })
            )
          )
        );
      return result;
    } finally {
      state.reservations.pop();
      state.capturing = previous;
      state.captureOrigin = previousOrigin;
    }
  };
  const prepareRetraction = (items: readonly ReviewFlagReceipt[], flagWrites = 0) => {
    healthy(state);
    if (!state.retractions)
      throw new ReviewFlagError('Flag retraction requires explicit protocol 3 attachment.');
    if (!Array.isArray(items) || items.length > MAX_REVIEW_FLAG_INTENTS)
      throw new ReviewFlagError('Invalid bounded flag retraction receipt.');
    if (!Number.isSafeInteger(flagWrites) || flagWrites < 0)
      throw new ReviewFlagError('Review ownership needs a complete bounded removal flag count.');
    if (doc.store.pendingStructs || doc.store.pendingDs)
      throw new ReviewFlagError('Review decisions must wait for complete room updates.');
    const withdrawn = new Set(
      [...state.retractions].map(([id, raw]) => decodeFlagRetraction(id, raw).id)
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
        throw new ReviewFlagError('Flag retraction receipt is stale or invalid.');
      validateFlagIntent(item.id, item.record);
      decodeFlagTarget(item.record, state.targets.get(item.record));
      ids.add(item.id);
      targets.add(item.record);
      encoded.push(JSON.stringify(item));
    }

    return { targets, encoded };
  };
  return Object.freeze({
    assertHealthy: () => {
      healthy(state);
      if (protocol !== '1' && state.targets.get('protocol') !== protocol)
        throw new ReviewFlagError('Unsupported review flag protocol.');
    },
    capture<T>(
      run: () => T,
      flagWrites: number,
      receipt?: (items: readonly ReviewFlagReceipt[]) => void
    ): T {
      return capture(run, flagWrites, receipt);
    },
    assertRetraction(items: readonly ReviewFlagReceipt[], flagWrites = 0): void {
      const { targets } = prepareRetraction(items, flagWrites);
      capture(() => {}, flagWrites, undefined, targets, items.length);
    },
    retractInTransaction<T>(items: readonly ReviewFlagReceipt[], run: () => T, flagWrites = 0): T {
      const { targets, encoded } = prepareRetraction(items, flagWrites);
      if (!doc._transaction?.local || doc._transaction.origin instanceof Y.UndoManager)
        throw new ReviewFlagError('Flag retraction requires an active local native transaction.');
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
        flagWrites,
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
      attached.delete(doc);
    },
  });
}
