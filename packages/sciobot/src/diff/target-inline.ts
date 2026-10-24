import { planImageReplacements } from './image-replacements';
import { mixedInlineAtoms } from './mixed-inline-atoms';
import { linkAwareBias } from './link-aware-text';
import { planRetainedAtomMarks } from './retained-atom-marks';
import { retainedAtomGaps } from './retained-atom-gaps';
import { withoutInlineAtoms } from './retained-atoms';
import { planInsertedAtoms } from './inserted-atoms';
import {
  findNode,
  type OoxmlNode,
  paragraphOffsetIndex,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import {
  compileInlineTarget,
  type CompiledInlineTarget,
  type InlineTargetMark,
} from './inline-target';
import type { LogicalParagraph } from './logical-block-types';
import type { BoundLogicalSource, BoundSourceBlock } from './bound-source-types';
import type { TargetSourceInput } from './target-source-types';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';
import { planParagraphText } from './paragraph-text';
import { planHyperlinkFormat } from './run-format';
import { markAt, planMarkDeltas, previewTextEdits, type MarkRun } from './insertion-marks';
import { hyperlinkCharacterStyle, inlineCodeCharacterStyles } from '../native/inline-code';
import { planLinkStructure } from './target-links';
import { planParagraphFormat } from './paragraph-format';
import { domainNodeFromAst } from './domain-ast';
import { planAtomicBlock } from './atomic-blocks';
import { targetAstKey, jsonDataKey } from './target-key';
import { withoutGeneratedCellSeparator } from './target-generated-separators';

export interface TargetParagraphSource {
  readonly input: TargetSourceInput;
  readonly source: BoundLogicalSource;
  readonly binding: BoundSourceBlock;
  readonly before: LogicalParagraph;
  readonly after: LogicalParagraph;
  /** Host resolves heading depths to actual native style IDs. Null refuses unresolved styles. */
  readonly nativeRange?: { readonly start: number; readonly end: number };
  readonly allocateBlockId?: () => string;
  readonly headingStyle?: (depth: number) => string | null;
  /** Internal anchored diff, computed before retained controls become synthetic text. */
  readonly anchoredText?: EditPlanResult;
  /** Internal source-to-target ownership, selected before image replacement changes keys. */
  readonly retainedAtomAlignment?: readonly number[];
  readonly textAnchors?: readonly { start: number; end: number }[];
}
const fail = (error: string): EditPlanResult => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
type Run = MarkRun;
const markKey = (marks: readonly InlineTargetMark[]) => marks.join('/');
const profile = (runs: readonly Run[]): Run[] => {
  const out: Run[] = [];
  for (const run of runs) {
    if (!run.value) continue;
    const previous = out.at(-1);
    if (previous && markKey(previous.marks) === markKey(run.marks)) previous.value += run.value;
    else out.push({ ...run });
  }
  return out;
};
const textRuns = (compiled: CompiledInlineTarget): Run[] =>
  compiled.tokens.flatMap((token) =>
    token.kind === 'text' ? [{ value: token.value, marks: token.marks }] : []
  );
function slice(runs: readonly Run[], start: number, end: number): Run[] {
  const out: Run[] = [];
  let offset = 0;
  for (const run of runs) {
    const a = Math.max(0, start - offset),
      b = Math.min(run.value.length, end - offset);
    if (a < b) out.push({ ...run, value: run.value.slice(a, b) });
    offset += run.value.length;
  }
  return out;
}
const ownHeader = (ast: TargetAstNode) =>
  Object.fromEntries(
    Object.entries(ast).filter(([key]) => !['type', 'depth', 'children', 'position'].includes(key))
  );
export function targetParagraphHtmlKey(paragraph: LogicalParagraph): string {
  const html = paragraph.html;
  const meaningful =
    html &&
    Object.keys(html).some(
      (key) =>
        !['type', 'name'].includes(key) &&
        !(key === 'attributes' && Array.isArray(html.attributes) && !html.attributes.length)
    );
  return meaningful ? targetAstKey({ ...html, type: 'mdxJsxFlowElement' }) : 'null';
}

/** Compose existing paragraph planners. Refuse unresolved inline intentions instead of dropping them. */
export function planTargetParagraph(context: TargetParagraphSource): EditPlanResult {
  const { input, source, binding, after } = context;
  const before = withoutGeneratedCellSeparator(input, source, binding, context.before);
  const tokenBefore = compileInlineTarget(before.ast),
    tokenAfter = compileInlineTarget(after.ast);
  if (!tokenBefore.success) return tokenBefore;
  if (!tokenAfter.success) return tokenAfter;
  const a = tokenBefore.target,
    b = tokenAfter.target,
    revision = source.index.revision,
    part = input.pkg.parts.get(source.index.partName)!;
  if (
    jsonDataKey(ownHeader(before.ast)) !== jsonDataKey(ownHeader(after.ast)) ||
    targetParagraphHtmlKey(before) !== targetParagraphHtmlKey(after)
  )
    return fail('Paragraph metadata needs an explicit native property mapping.');
  const ids = [...new Set(binding.contributions.map((entry) => entry.ownerId))];
  if (ids.length !== 1 || findNode(part, ids[0]!)?.kind !== 'paragraph')
    return fail('Changed logical paragraphs need one resolved native paragraph.');
  const paragraphId = ids[0]!,
    owner = source.index.owners.find((entry) => entry.id === paragraphId)!;
  const operations: CanonicalEdit[] = [];
  if (before.ast.type !== after.ast.type || before.ast.depth !== after.ast.depth) {
    const depth = after.ast.type === 'heading' ? Number(after.ast.depth) : 0;
    const style = depth ? context.headingStyle?.(depth) : null;
    if (depth && !style) return fail('The target heading needs a resolved native style ID.');
    const result = planParagraphFormat(
      { revision, part, paragraphId },
      style
        ? { properties: [{ localName: 'pStyle', attributes: { val: style } }] }
        : { properties: [], remove: ['pStyle'] }
    );
    if (!result.success) return result;
    operations.push(...result.plan.operations);
  }
  if (a.plainText !== null && b.plainText === null) {
    const skeleton = withoutInlineAtoms(after.ast);
    if (!skeleton) return fail('The retained paragraph has no resolved atom skeleton.');
    const text = planTargetParagraph({ ...context, after: { ...after, ast: skeleton } });
    if (!text.success) return text;
    const base = context.nativeRange?.start ?? owner.checklist?.native.end ?? 0;
    const atoms = planInsertedAtoms(
      {
        pkg: input.pkg,
        location: { kind: 'retained-paragraph', paragraphId },
        offset: base,
        allocateBlockId: context.allocateBlockId,
      },
      after.ast
    );
    if (!atoms.success) return atoms;
    const compiled = compileInlineTarget(skeleton);
    if (!compiled.success || compiled.target.plainText === null)
      return fail('The retained atom skeleton still has unresolved source.');
    const native = paragraphTextOf(part, paragraphId);
    if (native === null) return fail('The retained paragraph no longer exists.');
    const ownedLinks = new Set(
      atoms.plan.operations.flatMap((atom) =>
        atom.hyperlink ? [atom.hyperlink.path.join('/')] : []
      )
    );
    const spans = new Map<string, { start: number; end: number }>();
    let offset = base;
    for (const token of b.tokens) {
      if (token.kind !== 'text') continue;
      const key = token.hyperlink?.join('/');
      if (key !== undefined && ownedLinks.has(key)) {
        const previous = spans.get(key);
        spans.set(key, { start: previous?.start ?? offset, end: offset + token.value.length });
      }
      offset += token.value.length;
    }
    const atomLinks = b.hyperlinks.filter((link) => ownedLinks.has(link.path.join('/')));
    return createEditPlan(revision, [
      ...text.plan.operations.filter(
        (operation) =>
          operation.kind !== 'link' ||
          !atomLinks.some((link) => {
            const span = spans.get(link.path.join('/'));
            return (
              span?.start === operation.start &&
              span.end === operation.end &&
              jsonDataKey(link.target) === jsonDataKey(operation.target) &&
              link.title === operation.title
            );
          })
      ),
      {
        kind: 'inline-atoms',
        partName: part.name,
        paragraphId,
        offset: base,
        before: compiled.target.plainText,
        ast: after.ast,
        atoms: atoms.plan.operations,
      },
    ]);
  }
  if (a.plainText === null && b.plainText === null) {
    const images = planImageReplacements(context, part, paragraphId, a, b);
    if (images) {
      if ('success' in images) return images;
      const rest = planTargetParagraph({
        ...context,
        before: { ...before, ast: images.ast },
        retainedAtomAlignment: images.alignment,
      });
      return rest.success
        ? createEditPlan(revision, [...images.edits, ...rest.plan.operations])
        : rest;
    }
    const marks = planRetainedAtomMarks(
      context,
      part,
      paragraphId,
      a,
      b,
      context.nativeRange?.start ?? owner.checklist?.native.end ?? 0
    );
    if (marks) {
      if (!marks.success) return marks;
      return createEditPlan(revision, [...operations, ...marks.plan.operations]);
    }
  }
  if (a.plainText === null && !context.nativeRange) {
    const gaps = retainedAtomGaps(
      part,
      paragraphId,
      source,
      binding,
      a,
      b,
      owner.checklist?.native.end ?? 0
    );
    if (gaps) {
      // Right-to-left gaps keep every remaining source range in its original coordinates.
      for (const gap of [...gaps].reverse()) {
        const planned = planTargetParagraph({
          ...context,
          before: { ...before, ast: gap.before },
          after: { ...after, ast: gap.after },
          nativeRange: { start: gap.start, end: gap.end },
        });
        if (!planned.success) return planned;
        operations.push(...planned.plan.operations);
      }
      return createEditPlan(revision, operations);
    }
  }
  if (a.plainText === null && b.plainText === null) {
    const mixed = mixedInlineAtoms(
      context,
      part,
      paragraphId,
      a,
      b,
      owner.checklist?.native.end ?? 0
    );
    if (mixed)
      return planTargetParagraph({
        ...context,
        before: { ...before, ast: mixed.before },
        after: { ...after, ast: mixed.after },
        anchoredText: mixed.text,
        textAnchors: mixed.anchors,
      });
  }
  if (a.plainText === null || b.plainText === null) {
    const sameLinks =
      a.hyperlinks.length === b.hyperlinks.length &&
      a.hyperlinks.every((link, index) => {
        const next = b.hyperlinks[index]!;
        return jsonDataKey(link.target) === jsonDataKey(next.target) && link.title === next.title;
      });
    if (a.tokens.length !== b.tokens.length || !sameLinks)
      return fail('Changed opaque inline structure needs native atom reconciliation.');
    const blocks = buildBlockIndex(input.pkg);
    for (let i = 0; i < a.tokens.length; i++) {
      const old = a.tokens[i]!,
        next = b.tokens[i]!;
      const owner =
        old.hyperlink === null
          ? -1
          : a.hyperlinks.findIndex((link) => link.path.join('/') === old.hyperlink!.join('/'));
      const nextOwner =
        next.hyperlink === null
          ? -1
          : b.hyperlinks.findIndex((link) => link.path.join('/') === next.hyperlink!.join('/'));
      if (owner !== nextOwner)
        return fail('Opaque atom link ownership changes need native reconciliation.');
      if (old.kind !== next.kind || markKey(old.marks) !== markKey(next.marks))
        return fail('Opaque inline mark changes need native atom reconciliation.');
      if (old.kind === 'text' && next.kind === 'text') {
        if (old.value !== next.value)
          return fail('Text beside opaque atoms needs resolved native inline ranges.');
        continue;
      }
      if (old.kind !== 'atom' || next.kind !== 'atom') return fail('Opaque inline kinds differ.');
      if (targetAstKey(old.ast) === targetAstKey(next.ast)) continue;
      const position = old.ast.position as
        | { start?: { offset?: number }; end?: { offset?: number } }
        | undefined;
      const range =
        position?.start?.offset === undefined || position.end?.offset === undefined
          ? null
          : targetSourceRangeAt(source.parsed, position.start.offset, position.end.offset);
      const inline =
        range &&
        binding.inline.filter((item) => item.start === range.start && item.end === range.end);
      if (!inline || inline.length !== 1)
        return fail('The changed atom has no unique native control.');
      const block = blocks.byId.get(inline[0]!.blockId),
        decodedBefore = domainNodeFromAst(old.ast),
        decodedAfter = domainNodeFromAst(next.ast);
      if (
        !block ||
        block.payload.type !== 'source' ||
        decodedBefore?.type !== 'source' ||
        decodedAfter?.type !== 'source'
      )
        return fail('This atomic payload needs its application reconciliation codec.');
      const attributes = { ...block.payload.attributes };
      const names = new Set([...Object.keys(decodedBefore), ...Object.keys(decodedAfter)]);
      for (const name of names) {
        if (name === 'type') continue;
        if (
          jsonDataKey(decodedBefore[name] ?? null) === jsonDataKey(decodedAfter[name] ?? null) &&
          Object.hasOwn(decodedBefore, name) === Object.hasOwn(decodedAfter, name)
        )
          continue;
        if (Object.hasOwn(decodedAfter, name))
          Object.defineProperty(attributes, name, {
            value: decodedAfter[name],
            enumerable: true,
            configurable: true,
            writable: true,
          });
        else delete attributes[name];
      }
      const result = planAtomicBlock(
        { revision, pkg: input.pkg, blockId: block.id },
        { ...block.payload, attributes }
      );
      if (!result.success) return result;
      operations.push(...result.plan.operations);
    }
    return createEditPlan(revision, operations);
  }
  const task = owner.checklist;
  if (task && task.native.start !== 0)
    return fail('The task marker needs a contiguous native body range.');
  const base = context.nativeRange?.start ?? task?.native.end ?? 0;
  const beforeText = binding.generatedTask && !context.nativeRange ? '' : a.plainText;
  const afterText =
    binding.generatedTask &&
    after.list?.checked !== null &&
    after.list?.checked !== undefined &&
    b.plainText === '\u00a0'
      ? ''
      : b.plainText;
  const nativeText = paragraphTextOf(part, paragraphId);
  if (nativeText === null || nativeText.slice(base, context.nativeRange?.end) !== beforeText)
    return fail('The source paragraph needs complete visible-to-model text mapping.');
  const text =
    context.anchoredText ??
    planParagraphText(
      { revision, partName: part.name, paragraphId, text: beforeText, offset: base },
      afterText
    );
  if (!text.success) return text;
  const textEdits = linkAwareBias(
    part,
    paragraphId,
    base,
    text.plan.operations,
    b,
    context.textAnchors
  );
  const insertionMarks: CanonicalEdit[] = [];
  const runsBefore = binding.generatedTask ? [] : textRuns(a),
    runsAfter = afterText === '' && binding.generatedTask ? [] : textRuns(b);
  if (beforeText !== afterText) {
    let predicted = runsBefore;
    for (const edit of textEdits) {
      if (edit.kind !== 'text') return fail('Unexpected paragraph text operation.');
      const start = edit.start - base,
        end = edit.end - base;
      predicted = profile([
        ...slice(predicted, 0, start),
        {
          value: edit.after,
          marks: markAt(predicted, start, edit.bias ?? (edit.before ? 'right' : 'left')),
        },
        ...slice(predicted, end, Infinity),
      ]);
    }
    if (JSON.stringify(profile(predicted)) !== JSON.stringify(profile(runsAfter))) {
      // Inserted characters inherit their neighbor's marks. Format them afterwards, in post-text
      // coordinates, with old values read from the same native preview an executor will hold.
      const preview = previewTextEdits(part, textEdits);
      if (!preview.success) return preview;
      const marks = planMarkDeltas(
        {
          revision,
          part: preview.part,
          paragraphId,
          base,
          length: afterText.length,
          displayMode: input.projection.displayMode!,
          hyperlinkStyle: hyperlinkCharacterStyle(input.pkg),
          codeStyles: inlineCodeCharacterStyles(input.pkg),
        },
        predicted,
        runsAfter
      );
      if (!marks.success) return marks;
      insertionMarks.push(...marks.operations);
    }
  } else {
    const marks = planMarkDeltas(
      {
        revision,
        part,
        paragraphId,
        base,
        length: beforeText.length,
        displayMode: input.projection.displayMode!,
        hyperlinkStyle: hyperlinkCharacterStyle(input.pkg),
        codeStyles: inlineCodeCharacterStyles(input.pkg),
      },
      runsBefore,
      runsAfter
    );
    if (!marks.success) return marks;
    operations.push(...marks.operations);
  }
  // Same text and link spans: retarget the original wrappers in place.
  const sameLinks = (): EditPlanResult | readonly CanonicalEdit[] | null => {
    if (
      beforeText !== afterText ||
      a.hyperlinks.length !== b.hyperlinks.length ||
      a.hyperlinks.some((link, i) => link.title !== null && b.hyperlinks[i]!.title === null)
    )
      return null;
    const retargets: CanonicalEdit[] = [];
    const paragraph = findNode(part, paragraphId)!;
    const links: OoxmlNode[] = [],
      stack = [paragraph];
    while (stack.length) {
      const node = stack.pop()!;
      if (node.kind === 'hyperlink') links.push(node);
      else if (node.kind !== 'textValue') stack.push(...node.children);
    }
    const offsetIndex = paragraphOffsetIndex(
      paragraph as Extract<OoxmlNode, { kind: 'paragraph' }>
    );
    const ranges = (value: CompiledInlineTarget) => {
      const out = new Map<string, { start: number; end: number }>();
      let offset = base;
      for (const token of value.tokens) {
        if (token.kind !== 'text') continue;
        const end = offset + token.value.length;
        if (token.hyperlink) {
          const key = token.hyperlink.join('/'),
            old = out.get(key);
          out.set(key, { start: old?.start ?? offset, end });
        }
        offset = end;
      }
      return out;
    };
    const oldRanges = ranges(a),
      newRanges = ranges(b);
    for (let i = 0; i < a.hyperlinks.length; i++) {
      const old = a.hyperlinks[i]!,
        next = b.hyperlinks[i]!,
        range = oldRanges.get(old.path.join('/')),
        afterRange = newRanges.get(next.path.join('/'));
      if (!range || !afterRange || range.start !== afterRange.start || range.end !== afterRange.end)
        return null;
      const matches = links.filter((link) => {
        const span = offsetIndex.spanOf(link);
        return span?.start === range.start && span.end === range.end;
      });
      if (matches.length !== 1) return null;
      const result = planHyperlinkFormat(
        { revision, pkg: input.pkg, partName: part.name, hyperlinkId: matches[0]!.id },
        {
          ...(next.target.kind === 'external'
            ? { url: next.target.url }
            : { anchor: next.target.anchor }),
          tooltip: next.title,
        }
      );
      if (!result.success) return result;
      retargets.push(...result.plan.operations);
    }
    return retargets;
  };
  const linkOperations: CanonicalEdit[] = [];
  if (a.hyperlinks.length || b.hyperlinks.length) {
    const same = sameLinks();
    if (same && 'success' in same) return same;
    if (same) operations.push(...(same as readonly CanonicalEdit[]));
    else {
      // Link boundaries, insertion, removal, or text edits beside links: plan after the text edits.
      const preview = previewTextEdits(part, textEdits);
      if (!preview.success) return preview;
      const links = planLinkStructure(
        { revision, pkg: input.pkg, part: preview.part, paragraphId, base },
        b
      );
      if (!links.success) return links;
      linkOperations.push(...links.operations);
    }
  }
  operations.push(...textEdits, ...insertionMarks, ...linkOperations);
  return createEditPlan(revision, operations);
}
