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
type Piece = NonNullable<CanonicalTextNodeDescriptor['sourceText']>[number];
type Slice = Extract<Piece, { logicalId: string }>;
interface Copy {
  length: number;
  value: string | null;
  sources: TextSource[];
}
interface TextValue {
  readonly value: string | null;
}
export class TextSources {
  readonly enabled: boolean;
  private readonly values = new Map<string, TextSource[]>();
  private readonly objects = new WeakMap<object, readonly TextSource[]>();
  private readonly events = new Map<number, { id: string; slice: Slice; slot: number }[]>();
  private readonly targets = new Set<string>();
  private readonly copies = new Map<string, Copy[]>();
  private invalid = false;
  readonly unproven = new Set<string>();
  private readonly initial = new Set<string>();
  private readonly replayed = new Set<string>();
  private readonly replayedTargets = new Set<string>();
  isReplayedSource(id: string): boolean {
    return this.replayed.has(id);
  }
  constructor(effects: readonly CanonicalPrimitiveEffect[]) {
    // Typing retains fresh origins even beside metadata writes. Rebuilt text without
    // native source recipes still needs the established fragment reconciliation.
    const implicitText = effects.some(
      (effect) =>
        effect.kind === 'putNode' &&
        effect.descriptor.kind === 'textValue' &&
        effect.descriptor.sourceText === undefined
    );
    this.enabled = effects.some(
      (e) =>
        (!implicitText && e.kind === 'spliceText') ||
        (e.kind === 'putNode' &&
          e.descriptor.kind === 'textValue' &&
          e.descriptor.sourceText !== undefined)
    );
    if (!this.enabled) return;
    effects.forEach((effect, index) => {
      if (
        effect.kind !== 'putNode' ||
        effect.descriptor.kind !== 'textValue' ||
        effect.descriptor.sourceText === undefined
      )
        return;
      const pieces = effect.descriptor.sourceText;
      const id = effect.descriptor.logicalId;
      if (!Array.isArray(pieces) || pieces.length === 0 || rejectDangerousKey(id)) {
        this.invalid = true;
        return;
      }
      this.targets.add(id);
      const copies: Copy[] = [];
      this.copies.set(id, copies);
      pieces.forEach((piece, slot) => {
        if (
          !piece ||
          typeof piece !== 'object' ||
          !Number.isSafeInteger(piece.utf16Length) ||
          piece.utf16Length < 0
        ) {
          this.invalid = true;
          return;
        }
        copies.push({ value: null, sources: [], length: piece.utf16Length });
        if (!('logicalId' in piece)) return;
        const slice = piece as Slice;
        if (
          !Number.isSafeInteger(slice.beforeEffect) ||
          slice.beforeEffect < 0 ||
          slice.beforeEffect > index ||
          !Number.isSafeInteger(slice.utf16Start) ||
          slice.utf16Start < 0 ||
          typeof slice.logicalId !== 'string' ||
          !slice.logicalId ||
          rejectDangerousKey(slice.logicalId)
        ) {
          this.invalid = true;
          return;
        }
        const list = this.events.get(slice.beforeEffect) ?? [];
        list.push({ id, slice, slot });
        this.events.set(slice.beforeEffect, list);
      });
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
  finishSeed(): void {
    const knownSources = new Map<string, boolean>();
    for (const events of this.events.values())
      for (const { id, slice } of events) {
        const known = this.initial.has(id);
        knownSources.set(slice.logicalId, (knownSources.get(slice.logicalId) ?? true) && known);
        if (known) this.replayedTargets.add(id);
      }
    for (const [id, allKnown] of knownSources) if (allKnown) this.replayed.add(id);
  }
  isReplayedTarget(id: string): boolean {
    return this.replayedTargets.has(id);
  }
  get alreadyApplied(): boolean {
    return (
      !this.invalid &&
      this.targets.size > 0 &&
      [...this.targets].every((id) => this.initial.has(id))
    );
  }
  capture(index: number, nodeOf: (id: string) => TextValue | undefined): boolean {
    if (this.invalid) return false;
    const events = this.events.get(index) ?? [];
    for (const { id, slice, slot } of events) {
      // An already shown target belongs to a replay; its own native content is authoritative.
      if (this.values.has(id)) continue;
      const source = nodeOf(slice.logicalId);
      if (!source) {
        this.unproven.add(id);
        continue;
      }
      if (this.unproven.has(slice.logicalId)) this.unproven.add(id);
      if (source.value === null || slice.utf16Start + slice.utf16Length > source.value.length)
        return false;
      const sources = this.values.get(slice.logicalId);
      if (!sources || sources.length !== source.value.length) return false;
      if (slice.utf16Length === 0) continue;
      this.copies.get(id)![slot] = {
        length: slice.utf16Length,
        value: source.value.slice(slice.utf16Start, slice.utf16Start + slice.utf16Length),
        sources: sources.slice(slice.utf16Start, slice.utf16Start + slice.utf16Length),
      };
    }
    return true;
  }
  splice(id: string, value: string, start: number, remove: number, insert: string): boolean {
    if (!this.enabled) return true;
    const copied = this.copies.get(id);
    if (copied) {
      if (value !== '' || start !== 0 || remove !== 0) return false;
      if (copied.reduce((sum, piece) => sum + piece.length, 0) !== insert.length) return false;
      let offset = 0;
      const origins: TextSource[] = [];
      for (const piece of copied) {
        if (piece.value !== null && insert.slice(offset, offset + piece.length) !== piece.value)
          return false;
        offset += piece.length;
        for (let i = 0; i < piece.length; i++) origins.push(piece.sources[i] ?? null);
      }
      if (offset !== insert.length) return false;
      this.values.set(id, origins);
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
