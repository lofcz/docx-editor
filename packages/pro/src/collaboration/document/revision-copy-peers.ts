/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { attributeChanges } from './paragraph-text-diff.ts';
import { parseClientClock, type LogicalId } from './identity.ts';
import { itemIdentity } from './paragraph-text-identity.ts';
import { structAt } from './yjs-items.ts';
import type { ShownText, FollowingText } from './paragraph-text-follow.ts';
import type { RevisionCopyPosition } from './revision-text-copies.ts';

type Attributes = Readonly<Record<string, unknown>>;
function supersedes(attributes: Attributes, key: string): boolean {
  const raw = attributes.b;
  if (typeof raw !== 'string' || raw.length > 65536) return false;
  try {
    const value: unknown = JSON.parse(raw);
    return Array.isArray(value) && value.includes(key);
  } catch {
    return false;
  }
}
function branchAttributes(
  raw: Attributes,
  branch: RevisionCopyPosition,
  opposite: RevisionCopyPosition
): Attributes {
  const out = { ...raw };
  delete out[branch.role.key];
  delete out[opposite.role.key];
  out[branch.role.key] = branch.attributes[branch.role.key];
  if (typeof raw.t === 'string' && raw.t.length <= 65536) {
    try {
      const shell: unknown = JSON.parse(raw.t);
      if (shell && typeof shell === 'object' && !Array.isArray(shell))
        out.t = JSON.stringify({
          ...shell,
          k: branch.role.kind === 'del' ? 'deletedText' : 'text',
          l: branch.role.kind === 'del' ? 'delText' : 't',
        });
    } catch {
      /* The normal decoder handles an invalid source shell. */
    }
  }

  if (branch.attributes.w === undefined) delete out.w;
  else out.w = branch.attributes.w;
  if (typeof out.b === 'string') {
    try {
      const value: unknown = JSON.parse(out.b);
      if (Array.isArray(value)) {
        const remaining = value.filter(
          (entry) => entry !== branch.role.key && entry !== opposite.role.key
        );
        if (remaining.length) out.b = JSON.stringify(remaining);
        else delete out.b;
      }
    } catch {
      /* The normal decoder retains its existing malformed-data policy. */
    }
  }
  return out;
}
function typedAfter(text: Y.Text, shown: ShownText, at: number, left: number): boolean {
  let pointer = parseClientClock(shown.identities.origins[at] ?? '');
  for (let depth = 0; pointer && text.doc && depth < 64; depth++) {
    const item = structAt(text.doc, pointer);
    if (!item) return false;
    if (item.content instanceof Y.ContentString)
      return itemIdentity(text, item, pointer.clock - item.id.clock) === shown.identities.ids[left];
    pointer = item.origin;
  }
  return false;
}

/** A peer edits one live source; each opposite review snapshot projects that same edit once. */
export function revisionCopyPeers(
  text: Y.Text,
  shown: ShownText,
  paragraph: LogicalId,
  copies: ReadonlyMap<number, RevisionCopyPosition>
): ShownText {
  if (!copies.size) return shown;
  const attributes: Attributes[] = [],
    characters: boolean[] = [];
  for (const delta of text.toDelta()) {
    const width = typeof delta.insert === 'string' ? delta.insert.length : 1;
    for (let i = 0; i < width; i++) {
      attributes.push(delta.attributes ?? {});
      characters.push(typeof delta.insert === 'string');
    }
  }
  const overrides = new Map<number, Attributes>();
  const incoming: FollowingText[] = [];
  for (let start = 1; start < attributes.length; start++) {
    const left = copies.get(start - 1);
    if (!left || copies.has(start) || !characters[start] || shown.hidden.has(start)) continue;
    let end = start;
    while (end < attributes.length && !copies.has(end) && characters[end] && !shown.hidden.has(end))
      end++;
    const right = copies.get(end),
      other = copies.get(left.other),
      otherRight = right && copies.get(right.other);
    if (
      !right ||
      !other ||
      !otherRight ||
      left.role.key !== right.role.key ||
      other.role.key !== otherRight.role.key ||
      left.other >= right.other
    ) {
      start = end;
      continue;
    }
    const committed = supersedes(attributes[start]!, left.role.key);
    if (!committed && !(left.role.kind === 'del' && typedAfter(text, shown, start, start - 1))) {
      start = end;
      continue;
    }
    let at = start;
    while (at < end) {
      const raw = attributes[at]!;
      let until = at + 1;
      while (until < end && attributes[until] === raw) until++;
      const source = branchAttributes(raw, left, other),
        target = branchAttributes(raw, other, left);
      const positions: number[] = [];
      for (let pos = at; pos < until; pos++) {
        overrides.set(pos, source);
        positions.push(pos);
      }
      incoming.push({
        source: paragraph,
        after: shown.identities.ids[left.other]!,
        before: false,
        anchorItem: shown.identities.items[left.other]!,
        positions,
        identities: positions.map((pos) => shown.identities.ids[pos]!),
        items: positions.map((pos) => shown.identities.items[pos]!),
        attributesOverride: target,
      });
      at = until;
    }
    start = end;
  }
  return overrides.size
    ? { ...shown, attributeOverrides: overrides, incoming: [...shown.incoming, ...incoming] }
    : shown;
}

/** Persist the source snapshot membership seen by the native decision preflight. */
export function persistRevisionCopyMembership(
  text: Y.Text,
  overrides: ReadonlyMap<number, Attributes>
): void {
  const raw: Attributes[] = [];
  for (const entry of text.toDelta()) {
    const width = typeof entry.insert === 'string' ? entry.insert.length : 1;
    for (let at = 0; at < width; at++) raw.push(entry.attributes ?? {});
  }
  const pending: { position: number; changes: Record<string, string | null>; key: string }[] = [];
  for (const [position, override] of overrides) {
    const desired: Record<string, string> = {};
    for (const [key, value] of Object.entries(override)) {
      if (typeof value === 'string') desired[key] = value;
    }
    const changes = attributeChanges(
      Object.fromEntries(
        Object.entries(raw[position] ?? {}).filter(
          (entry): entry is [string, string] => typeof entry[1] === 'string'
        )
      ),
      desired
    );
    if (changes) pending.push({ position, changes, key: JSON.stringify(changes) });
  }
  pending.sort((a, b) => b.position - a.position);
  for (let index = 0; index < pending.length; ) {
    const first = pending[index]!;
    let start = first.position;
    let next = index + 1;
    while (
      next < pending.length &&
      pending[next]!.position === start - 1 &&
      pending[next]!.key === first.key
    ) {
      start -= 1;
      next += 1;
    }
    text.format(start, first.position - start + 1, first.changes);
    index = next;
  }
}
