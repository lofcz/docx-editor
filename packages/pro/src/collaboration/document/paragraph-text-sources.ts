/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/** Transaction-local identities for native text slices, before shared characters are written. */
import type { OoxmlNode } from '@docx-editor.dev/core/store';
import type {
  CanonicalPrimitiveEffect,
  CanonicalTextNodeDescriptor,
} from '@docx-editor.dev/core/collaboration/replication';
import { rejectDangerousKey } from './limits.ts';

export type TextSource = string | null;
type Slice = NonNullable<CanonicalTextNodeDescriptor['sourceText']>;
interface TextValue {
  readonly value: string | null;
}
export class TextSources {
  readonly enabled: boolean;
  private readonly values = new Map<string, TextSource[]>();
  private readonly objects = new WeakMap<object, readonly TextSource[]>();
  private readonly events = new Map<number, { id: string; slice: Slice }[]>();
  private readonly copies = new Map<string, { value: string; sources: TextSource[] }>();
  private invalid = false;
  private readonly initial = new Set<string>();
  private readonly replayed = new Set<string>();
  isReplayedSource(id: string): boolean {
    return this.replayed.has(id);
  }
  constructor(effects: readonly CanonicalPrimitiveEffect[]) {
    this.enabled = effects.some(
      (e) =>
        e.kind === 'putNode' &&
        e.descriptor.kind === 'textValue' &&
        e.descriptor.sourceText !== undefined
    );
    if (!this.enabled) return;
    effects.forEach((effect, index) => {
      if (
        effect.kind !== 'putNode' ||
        effect.descriptor.kind !== 'textValue' ||
        !effect.descriptor.sourceText
      )
        return;
      const slice = effect.descriptor.sourceText;
      if (
        !Number.isSafeInteger(slice.beforeEffect) ||
        slice.beforeEffect < 0 ||
        slice.beforeEffect > index ||
        !Number.isSafeInteger(slice.utf16Start) ||
        slice.utf16Start < 0 ||
        !Number.isSafeInteger(slice.utf16Length) ||
        slice.utf16Length < 0 ||
        typeof slice.logicalId !== 'string' ||
        !slice.logicalId ||
        rejectDangerousKey(slice.logicalId) ||
        rejectDangerousKey(effect.descriptor.logicalId)
      ) {
        this.invalid = true;
        return;
      }
      const list = this.events.get(slice.beforeEffect) ?? [];
      list.push({ id: effect.descriptor.logicalId, slice });
      this.events.set(slice.beforeEffect, list);
    });
  }
  seed(id: string, value: string, sources: readonly TextSource[]): void {
    if (this.enabled) {
      this.initial.add(id);
      if (sources.length !== value.length) {
        this.invalid = true;
        return;
      }
      this.values.set(id, [...sources]);
    }
  }
  capture(index: number, nodeOf: (id: string) => TextValue | undefined): boolean {
    if (this.invalid) return false;
    const events = this.events.get(index) ?? [];
    const knownSources = new Map<string, boolean>();
    for (const { id, slice } of events)
      knownSources.set(
        slice.logicalId,
        (knownSources.get(slice.logicalId) ?? true) && this.initial.has(id)
      );
    // Fresh target identities already present identify a repeated native split.
    for (const [id, allKnown] of knownSources) if (allKnown) this.replayed.add(id);
    for (const { id, slice } of events) {
      // An already shown target belongs to a replay; its own native content is authoritative.
      if (this.values.has(id)) continue;
      const source = nodeOf(slice.logicalId);
      if (!source) continue; // Record-owned opaque content has no paragraph-text provenance.
      if (source.value === null || slice.utf16Start + slice.utf16Length > source.value.length)
        return false;
      const sources = this.values.get(slice.logicalId);
      if (!sources || sources.length !== source.value.length || this.copies.has(id)) return false;
      this.copies.set(id, {
        value: source.value.slice(slice.utf16Start, slice.utf16Start + slice.utf16Length),
        sources: sources.slice(slice.utf16Start, slice.utf16Start + slice.utf16Length),
      });
    }
    return true;
  }
  splice(id: string, value: string, start: number, remove: number, insert: string): boolean {
    if (!this.enabled) return true;
    const copied = this.copies.get(id);
    if (copied) {
      if (value !== '' || start !== 0 || remove !== 0 || copied.value !== insert) return false;
      this.values.set(id, copied.sources);
      this.copies.delete(id);
      return true;
    }
    const sources = this.values.get(id) ?? Array<TextSource>(value.length).fill(null);
    if (sources.length !== value.length) return false;
    this.values.set(id, [
      ...sources.slice(0, start),
      ...Array<TextSource>(insert.length).fill(null),
      ...sources.slice(start + remove),
    ]);
    return true;
  }
  remember(node: OoxmlNode, id: string): void {
    if (!this.enabled || node.kind !== 'textValue') return;
    this.objects.set(node, this.values.get(id) ?? Array<TextSource>(node.value.length).fill(null));
  }
  readonly of = (node: OoxmlNode): readonly TextSource[] | undefined => this.objects.get(node);
}
