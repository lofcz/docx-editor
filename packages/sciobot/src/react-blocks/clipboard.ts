import {
  beginHistoryGroupFrames,
  type DocxEditorInstance,
  type PaginatedSurface,
} from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import {
  blockIdFromTag,
  blockTag,
  readPayloadStore,
  withPayloadRecord,
} from '../blocks/payload-opc-codec';
import {
  decodeBlockPayload,
  MAX_BLOCK_ID_LENGTH,
  type BlockPayload,
  type JsonValue,
} from '../blocks/payload-schema';
import { readBlockSelection } from './selection';
import { isWidgetInputEvent } from './input-boundary';

/** Private clipboard flavour. Native Word and HTML flavours carry the readable fallback. */
export const SCIOBOT_BLOCKS_MIME = 'application/x-sciobot-blocks+json';
const MAX_TRANSFER_BYTES = 4 * 1024 * 1024;
const MAX_TRANSFER_BLOCKS = 2000;

export interface BlockClipboardEntry {
  readonly id: string;
  readonly label: string;
  readonly payload: BlockPayload;
}
export interface BlockClipboardData {
  readonly version: 1;
  readonly blocks: readonly BlockClipboardEntry[];
}
export interface BlockPasteResult {
  /** Old block ID to the fresh ID assigned in this document. */
  readonly remapped: ReadonlyMap<string, string>;
  /** Pasted block tags without a known payload. Their controls keep native content only. */
  readonly untagged: readonly string[];
}
export interface BlockClipboardOptions {
  /** Fresh block IDs. Must return valid, unused IDs. */
  readonly createId?: (old: BlockPayload | null, used: ReadonlySet<string>) => string;
  readonly onPaste?: (result: BlockPasteResult) => void;
  readonly onError?: (error: Error) => void;
}
export interface BlockClipboard {
  dispose(): void;
}

interface TaggedControl {
  readonly controlId: string;
  readonly id: string;
  readonly partName: string;
}

function taggedControls(pkg: OoxmlPackage): TaggedControl[] {
  const out: TaggedControl[] = [];
  for (const [partName, part] of pkg.parts) {
    if (!isStoryPart(part)) continue;
    for (const { node } of contentControlsIn(part.root)) {
      const id = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
      if (id !== null) out.push({ controlId: node.id, id, partName });
    }
  }
  return out;
}

/** Bounded JSON reader for private clipboard data. Invalid data reads as absent. */
export function readBlockClipboardData(raw: string | null | undefined): BlockClipboardData | null {
  if (!raw || raw.length > MAX_TRANSFER_BYTES) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) return null;
  const record = parsed as { version?: unknown; blocks?: unknown };
  if (record.version !== 1 || !Array.isArray(record.blocks)) return null;
  if (record.blocks.length > MAX_TRANSFER_BLOCKS) return null;
  const blocks: BlockClipboardEntry[] = [];
  for (const entry of record.blocks as unknown[]) {
    if (typeof entry !== 'object' || entry === null) return null;
    const { label, payload } = entry as { label?: unknown; payload?: unknown };
    if (typeof label !== 'string' || label.length > 4096) return null;
    const decoded = decodeBlockPayload(JSON.stringify(payload ?? null));
    if (decoded.status !== 'supported') return null;
    blocks.push(Object.freeze({ id: decoded.value.id, label, payload: decoded.value }));
  }
  return Object.freeze({ version: 1, blocks: Object.freeze(blocks) });
}

/** Payloads for every registered block the current selection reaches. */
export function collectBlockClipboardData(editor: DocxEditorInstance): BlockClipboardData | null {
  const surface = editor.surface;
  const selection = readBlockSelection(editor);
  if (!surface || !selection || selection.collapsed || !selection.blocks.length) return null;
  const pkg = surface.session.currentPackage();
  const index = buildBlockIndex(pkg);
  const store = readPayloadStore(pkg);
  const blocks: BlockClipboardEntry[] = [];
  const add = (id: string) => {
    const block = index.byId.get(id);
    if (!block || blocks.some((entry) => entry.id === id)) return;
    const label = store.status === 'ready' ? (store.entries.get(id)?.label ?? id) : id;
    blocks.push({ id, label, payload: block.payload });
  };
  for (const entry of selection.blocks) add(entry.id);
  // A copied container carries its registered descendants, even outside the range ends.
  for (const block of index.byId.values()) {
    let parent = block.parentId;
    while (parent !== null) {
      if (blocks.some((entry) => entry.id === parent)) {
        add(block.id);
        break;
      }
      parent = index.byId.get(parent)?.parentId ?? null;
    }
  }
  return blocks.length ? { version: 1, blocks } : null;
}

function defaultId(old: BlockPayload | null, used: ReadonlySet<string>): string {
  const base = (old?.type ?? 'block').replace(/[^A-Za-z0-9_-]/g, '-').replace(/^[^A-Za-z0-9]+/, '');
  for (;;) {
    const random = Array.from(globalThis.crypto.getRandomValues(new Uint8Array(6)), (byte) =>
      byte.toString(36).padStart(2, '0')
    ).join('');
    const id = `${(base || 'block').slice(0, MAX_BLOCK_ID_LENGTH - random.length - 1)}-${random}`;
    if (!used.has(id)) return id;
  }
}

function remapValue(value: JsonValue, ids: ReadonlyMap<string, string>): JsonValue {
  if (typeof value === 'string') return ids.get(value) ?? value;
  if (Array.isArray(value)) return value.map((item) => remapValue(item, ids));
  if (value && typeof value === 'object') {
    const out: Record<string, JsonValue> = {};
    for (const [key, item] of Object.entries(value)) out[key] = remapValue(item, ids);
    return out;
  }
  return value;
}

/**
 * Give blocks that a paste just created fresh IDs and payload records.
 * Internal references between pasted blocks follow the new IDs. Existing blocks stay intact.
 */
export function repairPastedBlocks(
  editor: DocxEditorInstance,
  before: ReadonlySet<string>,
  data: BlockClipboardData | null,
  options: BlockClipboardOptions = {}
): BlockPasteResult | null {
  const surface = editor.surface;
  if (!surface) return null;
  surface.flushPendingInput();
  const pkg = surface.session.currentPackage();
  const controls = taggedControls(pkg);
  const pasted = controls.filter((control) => !before.has(control.controlId));
  if (!pasted.length) return null;
  const store = readPayloadStore(pkg);
  const known = new Map<string, BlockClipboardEntry>();
  for (const entry of data?.blocks ?? []) known.set(entry.id, entry);
  // A same-document copy without private data can still use the stored payload.
  if (store.status === 'ready')
    for (const [id, stored] of store.entries)
      if (!known.has(id) && stored.decoded.status === 'supported')
        known.set(id, { id, label: stored.label, payload: stored.decoded.value });
  const used = new Set<string>(controls.map((control) => control.id));
  if (store.status === 'ready') for (const id of store.entries.keys()) used.add(id);
  const remapped = new Map<string, string>();
  const untagged: string[] = [];
  for (const control of pasted) {
    if (remapped.has(control.id) || untagged.includes(control.id)) continue;
    const source = known.get(control.id);
    if (!source) {
      untagged.push(control.id);
      continue;
    }
    const id = (options.createId ?? defaultId)(source.payload, used);
    if (used.has(id) || blockIdFromTag(blockTag(id)) !== id)
      throw new Error('Block clipboard created an invalid or duplicate ID.');
    used.add(id);
    remapped.set(control.id, id);
  }
  const ops: TreeDocOp[] = pasted.map((control) => {
    const id = remapped.get(control.id);
    return {
      op: 'setContentControlProperties',
      controlId: control.controlId,
      tag: id === undefined ? '' : blockTag(id),
    } as TreeDocOp;
  });
  const records = [...remapped].map(([old, id]) => {
    const source = known.get(old)!;
    const payload = remapValue(source.payload, remapped) as BlockPayload;
    return { payload: { ...payload, id }, label: source.label };
  });
  const result = surface.applyAutomationOps(() => ops, { kind: 'body' }, [
    (working) => {
      let next = working;
      for (const record of records) {
        const written = withPayloadRecord(next, record.payload, record.label);
        if (!written.ok) throw new Error(`Pasted block payload refused: ${written.reason}`);
        next = written.pkg;
      }
      return next;
    },
  ]);
  if (result.rejected || !result.committed)
    throw new Error(String(result.reason ?? 'The pasted block repair did not commit.'));
  return { remapped, untagged };
}

/**
 * Paste native clipboard flavours, then repair the pasted block tags in the same undo step.
 * Answers whether the native paste changed the document.
 */
export function pasteWithBlocks(
  editor: DocxEditorInstance,
  transfer: {
    readonly text: string;
    readonly html: string | null;
    readonly data: BlockClipboardData | null;
  },
  options: BlockClipboardOptions = {}
): boolean {
  const surface = editor.surface;
  if (!surface) return false;
  surface.flushPendingInput();
  const before = new Set(taggedControls(surface.session.currentPackage()).map((c) => c.controlId));
  const group = editor.beginHistoryGroup();
  const frames = beginHistoryGroupFrames(editor, group);
  if (!frames) throw new Error('The block paste history group is not open.');
  try {
    const changed = frames.run(() => surface.pasteRich(transfer.text, transfer.html || null));
    if (!changed) return false;
    const result = frames.run(() => repairPastedBlocks(editor, before, transfer.data, options));
    if (result) options.onPaste?.(result);
    return true;
  } finally {
    frames.finish();
    group.end();
  }
}

interface ClipboardLifetime {
  readonly clients: Map<symbol, BlockClipboardOptions>;
  readonly stop: () => void;
}
const clipboards = new WeakMap<PaginatedSurface, ClipboardLifetime>();

/** Retain one clipboard handler per surface. Independent callers release their own lease. */
export function mountBlockClipboard(
  editor: DocxEditorInstance,
  options: BlockClipboardOptions = {}
): BlockClipboard {
  const surface = editor.surface;
  if (!surface) throw new Error('Attach the editor before mounting the block clipboard.');
  let lifetime = clipboards.get(surface);
  if (
    lifetime &&
    options.createId &&
    [...lifetime.clients.values()].some(
      (client) => client.createId && client.createId !== options.createId
    )
  )
    throw new Error('The surface clipboard already has a different block ID factory.');
  const token = Symbol();
  if (!lifetime) {
    const clients = new Map([[token, options]]);
    const shared: BlockClipboardOptions = {
      createId: (old, used) =>
        ([...clients.values()].find((client) => client.createId)?.createId ?? defaultId)(old, used),
      onPaste: (result) => {
        for (const client of [...clients.values()]) client.onPaste?.(result);
      },
      onError: (error) => {
        for (const client of [...clients.values()]) client.onError?.(error);
      },
    };
    const handle = attachBlockClipboard(editor, shared);
    lifetime = { clients, stop: () => handle.dispose() };
    clipboards.set(surface, lifetime);
  } else lifetime.clients.set(token, options);
  const retained = lifetime;
  return {
    dispose() {
      if (!retained.clients.delete(token) || retained.clients.size) return;
      clipboards.delete(surface);
      retained.stop();
    },
  };
}

/** Carry registered block payloads through native clipboard events. */
function attachBlockClipboard(
  editor: DocxEditorInstance,
  options: BlockClipboardOptions
): BlockClipboard {
  const surface = editor.surface;
  if (!surface) throw new Error('Attach the editor before mounting the block clipboard.');
  let host = surface.hostElement();
  const fail = (error: unknown) =>
    options.onError?.(error instanceof Error ? error : new Error(String(error)));
  const write = (event: Event) => {
    const transfer = (event as ClipboardEvent).clipboardData;
    if (!transfer) return;
    try {
      const data = collectBlockClipboardData(editor);
      if (data) transfer.setData(SCIOBOT_BLOCKS_MIME, JSON.stringify(data));
    } catch (error) {
      fail(error);
    }
  };
  // Bubble phase: the native handler has already written its text and HTML flavours.
  const onCopy = (event: Event) => {
    if (isWidgetInputEvent(event)) return;
    if (event.defaultPrevented) write(event);
  };
  // Capture phase for cut: the native handler deletes the selection after writing flavours.
  const onCut = (event: Event) => {
    if (!isWidgetInputEvent(event)) write(event);
  };
  // Capture phase. Native fragments can carry block tags, so their paste and the tag repair
  // run as two writes of one explicit undo step. Other pastes keep the native handler.
  const onPaste = (event: Event) => {
    if (isWidgetInputEvent(event)) return;
    const transfer = (event as ClipboardEvent).clipboardData;
    if (!transfer) return;
    const html = transfer.getData('text/html');
    const raw = transfer.getData(SCIOBOT_BLOCKS_MIME);
    if (!raw && !html.includes('data-docx-fragment')) return;
    const text = transfer.getData('text/plain');
    if (!text && !html) return;
    event.preventDefault();
    event.stopPropagation();
    try {
      pasteWithBlocks(editor, { text, html, data: readBlockClipboardData(raw) }, options);
    } catch (error) {
      fail(error);
    }
  };
  const listen = () => {
    host.addEventListener('copy', onCopy);
    host.addEventListener('cut', onCut, true);
    host.addEventListener('paste', onPaste, true);
  };
  const unbind = () => {
    host.removeEventListener('copy', onCopy);
    host.removeEventListener('cut', onCut, true);
    host.removeEventListener('paste', onPaste, true);
  };
  const bind = () => {
    const next = surface.hostElement();
    if (next === host) return;
    unbind();
    host = next;
    listen();
  };
  listen();
  const stopHost = editor.on('selectionChange', bind);
  return {
    dispose() {
      stopHost();
      unbind();
    },
  };
}
