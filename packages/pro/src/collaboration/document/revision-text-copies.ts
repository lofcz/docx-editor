/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import { readRevisions } from './paragraph-text-revisions.ts';
import type { DocumentLimits } from './limits.ts';
import type { Token } from './paragraph-text-diff.ts';
import type { TextIdentities } from './paragraph-text-identity.ts';

type Attributes = Readonly<Record<string, unknown>>;
export interface RevisionCopyRole {
  readonly kind: 'ins' | 'del';
  readonly key: string;
  readonly author: string;
  readonly date: string | null;
}
function role(
  attributes: Attributes,
  limits: DocumentLimits,
  paragraph: string
): RevisionCopyRole | null {
  const revisions = readRevisions(attributes, limits, paragraph);
  if (revisions.length !== 1) return null;
  const revision = revisions[0]!,
    node = revision.shell;
  if (node.namespaceUri !== WML_NAMESPACE_URI || !['ins', 'del'].includes(node.localName))
    return null;
  const attribute = (name: string) =>
    node.attributes.find(
      (entry) => entry.namespaceUri === WML_NAMESPACE_URI && entry.localName === name
    )?.value;
  const author = attribute('author');
  return author
    ? {
        kind: node.localName as 'ins' | 'del',
        key: revision.key,
        author,
        date: attribute('date') ?? null,
      }
    : null;
}
function paired(a: RevisionCopyRole | null, b: RevisionCopyRole | null): boolean {
  return !!a && !!b && a.kind !== b.kind && a.author === b.author && a.date === b.date;
}

/** Only an explicit native source shared by opposite review copies can retain a second identity. */
export function pairedRevisionSources(
  writes: readonly { readonly id: string; readonly after: readonly Token[] }[],
  limits: DocumentLimits
): ReadonlySet<string> {
  const result = new Set<string>();
  for (const write of writes) {
    const copies = new Map<string, Token[]>();
    for (const token of write.after) {
      if (!token.source || typeof token.insert !== 'string') continue;
      const entries = copies.get(token.source) ?? [];
      if (entries.length < 3) entries.push(token);
      copies.set(token.source, entries);
    }
    for (const [source, entries] of copies) {
      if (entries.length !== 2 || entries[0]!.insert !== entries[1]!.insert) continue;
      if (
        paired(
          role(entries[0]!.attributes, limits, write.id),
          role(entries[1]!.attributes, limits, write.id)
        )
      )
        result.add(source);
    }
  }
  return result;
}

export interface RevisionCopyPosition {
  readonly other: number;
  readonly role: RevisionCopyRole;
  readonly attributes: Attributes;
}
/** Read paired snapshots through persisted character identities; equal text never establishes a pair. */
export function revisionCopyPositions(
  text: Y.Text,
  ids: TextIdentities,
  limits: DocumentLimits,
  paragraph: string
): ReadonlyMap<number, RevisionCopyPosition> {
  const delta = text.toDelta() as { insert: unknown; attributes?: Record<string, unknown> }[];
  if (
    !delta.some((entry) => Object.keys(entry.attributes ?? {}).some((key) => key.startsWith('v:')))
  )
    return new Map();
  const groups = new Map<string, number[]>();
  ids.ids.forEach((identity, position) => {
    if (!identity) return;
    const entries = groups.get(identity) ?? [];
    if (entries.length < 3) entries.push(position);
    groups.set(identity, entries);
  });
  const candidates = new Set<number>();
  for (const entries of groups.values())
    if (entries.length === 2) for (const at of entries) candidates.add(at);
  const attributes = new Map<number, Attributes>();
  let at = 0;
  for (const entry of delta) {
    const width = typeof entry.insert === 'string' ? entry.insert.length : 1;
    if (typeof entry.insert === 'string')
      for (let offset = 0; offset < width; offset++)
        if (candidates.has(at + offset)) attributes.set(at + offset, entry.attributes ?? {});
    at += width;
  }
  const result = new Map<number, RevisionCopyPosition>();
  for (const entries of groups.values()) {
    if (entries.length !== 2) continue;
    const [left, right] = entries as [number, number];
    const a = attributes.get(left),
      b = attributes.get(right);
    if (!a || !b) continue;
    const ra = role(a, limits, paragraph),
      rb = role(b, limits, paragraph);
    if (!paired(ra, rb)) continue;
    result.set(left, { other: right, role: ra!, attributes: a });
    result.set(right, { other: left, role: rb!, attributes: b });
  }
  return result;
}
