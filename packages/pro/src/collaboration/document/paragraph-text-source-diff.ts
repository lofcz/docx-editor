/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/** Exact source identities decide retention; inserted labels cannot borrow equal old letters. */
import { diffTokens, type Step, type Token } from './paragraph-text-diff.ts';

interface Unit {
  readonly token: Token;
  readonly positions: readonly number[];
}
function units(tokens: readonly Token[], side: string): Unit[] {
  const out: Unit[] = [];
  for (let index = 0; index < tokens.length; index++) {
    const first = tokens[index]!;
    const positions = [index];
    if (typeof first.insert !== 'string') {
      out.push({ token: first, positions });
      continue;
    }
    const code = first.insert.charCodeAt(0);
    const next = tokens[index + 1];
    if (
      code >= 0xd800 &&
      code <= 0xdbff &&
      typeof next?.insert === 'string' &&
      (first.position === undefined || next.position === first.position + 1)
    ) {
      const low = next.insert.charCodeAt(0);
      if (low >= 0xdc00 && low <= 0xdfff) positions.push(++index);
    }
    const characters = positions.map((position) => tokens[position]!.insert).join('');
    const sources = positions.map((position) => tokens[position]!.source);
    const retained = sources.every((source) => typeof source === 'string');
    // Synthetic embeds keep the ordinary diff's weak text heuristics out of this path.
    const key = retained ? `s${JSON.stringify([characters, sources])}` : `n${side}:${index}`;
    out.push({ positions, token: { ...first, key, strong: key, insert: { n: key } } });
  }
  return out;
}

export function diffSourceTokens(before: readonly Token[], after: readonly Token[]): Step[] {
  const left = units(before, 'before');
  const right = units(after, 'after');
  const out: Step[] = [];
  for (const step of diffTokens(
    left.map((unit) => unit.token),
    right.map((unit) => unit.token)
  )) {
    if (step.op === 'del') {
      for (const position of left[step.before]!.positions)
        out.push({ op: 'del', before: position });
    } else if (step.op === 'ins') {
      for (const position of right[step.after]!.positions) out.push({ op: 'ins', after: position });
    } else {
      const previous = left[step.before]!.positions;
      const next = right[step.after]!.positions;
      previous.forEach((position, index) =>
        out.push({ op: 'eq', before: position, after: next[index]! })
      );
    }
  }
  return out;
}
