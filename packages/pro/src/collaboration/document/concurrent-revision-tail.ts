/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type * as Y from 'yjs';
import { WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import { isNextClock, parseClientClock, type LogicalId } from './identity.ts';
import type { DocumentLimits } from './limits.ts';
import { readRevisions } from './paragraph-text-revisions.ts';
import type { ShownText } from './paragraph-text-follow.ts';
import { itemIdentity, type TextIdentities } from './paragraph-text-identity.ts';
import { structAt } from './yjs-items.ts';

/**
 * Read the original typing boundary through formatting clocks, including deleted clocks.
 * A restored source character can still show under its original identity. Testing identity
 * membership before reading content also handles markers collected differently by peers.
 * The walk counts clocks, not surviving structs, so collection cannot change its bound.
 */
function authoredBoundary(
  text: Y.Text,
  ids: TextIdentities,
  at: number,
  visible: ReadonlySet<string | null>
): string | null {
  const original = parseClientClock(ids.ids[at] ?? '');
  const doc = text.doc;
  if (!original || !doc) return null;
  const source = structAt(doc, original);
  if (!source) return null;
  let pointer =
    original.clock === source.id.clock
      ? source.origin
      : { client: original.client, clock: original.clock - 1 };
  for (let depth = 0; pointer && depth < 64; depth += 1) {
    const item = structAt(doc, pointer);
    if (!item) return null;
    const identity = itemIdentity(text, item, pointer.clock - item.id.clock);
    if (visible.has(identity)) return identity;
    pointer =
      pointer.clock === item.id.clock
        ? item.origin
        : { client: pointer.client, clock: pointer.clock - 1 };
  }
  return null;
}

/** Keep a tracked terminal append before text concurrently typed at its source boundary. */
export function concurrentRevisionTail(
  text: Y.Text,
  shown: ShownText,
  paragraphId: LogicalId,
  limits: DocumentLimits
): ShownText {
  if (shown.hidden.size || shown.incoming.length) return shown;
  const ids = shown.identities;
  const length = ids.ids.length;
  if (length < 3 || length > limits.maxTextLength) return shown;
  let start = length - 1;
  while (start > 0 && isNextClock(ids.ids[start - 1]!, ids.ids[start]!)) start -= 1;
  if (!ids.starts[start] || ids.moved[start] || start === 0) return shown;
  let peer = start - 1;
  while (peer > 0 && isNextClock(ids.ids[peer - 1]!, ids.ids[peer]!)) peer -= 1;
  if (!ids.starts[peer] || ids.clients[peer] === ids.clients[start]) return shown;
  const visible = new Set(ids.ids);
  const origin = authoredBoundary(text, ids, start, visible);
  if (!origin || authoredBoundary(text, ids, peer, visible) !== origin) return shown;
  if (ids.ids.indexOf(origin) + 1 !== peer) return shown;
  const revisions: ReturnType<typeof readRevisions>[] = [];
  const characters: boolean[] = [];
  for (const delta of text.toDelta()) {
    const count = typeof delta.insert === 'string' ? delta.insert.length : 1;
    const entries = readRevisions(delta.attributes ?? {}, limits, paragraphId);
    for (let i = 0; i < count; i += 1) {
      revisions.push(entries);
      characters.push(typeof delta.insert === 'string');
    }
  }
  const revision = revisions[start]?.[0];
  if (
    !revision ||
    revision.shell.namespaceUri !== WML_NAMESPACE_URI ||
    revision.shell.localName !== 'ins'
  )
    return shown;
  for (let at = peer; at < start; at += 1)
    if (!characters[at] || ids.moved[at] || revisions[at]?.length) return shown;
  const positions: number[] = [];
  for (let at = start; at < length; at += 1) {
    if (
      !characters[at] ||
      revisions[at]?.length !== 1 ||
      revisions[at]?.[0]?.signature !== revision.signature
    )
      return shown;
    positions.push(at);
  }
  return {
    ...shown,
    hidden: new Set(positions),
    incoming: [
      {
        source: paragraphId,
        after: origin,
        before: false,
        positions,
        identities: positions.map((at) => ids.ids[at]!),
        items: positions.map((at) => ids.items[at]!),
      },
    ],
  };
}
