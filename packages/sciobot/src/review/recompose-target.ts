import { preserveUnchangedTableSyntax } from './recompose-table-syntax';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  applyTreeOp,
  collectRevisionSites,
  headlessViewOfStore,
  TreePackageStore,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { previewDirectPlan } from '../ai/apply-direct';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { normalizeLogicalBlocks, type LogicalBlock } from '../diff/logical-blocks';
import { planTargetDocument, type TargetDocumentOptions } from '../diff/plan-target';
import { projectMdx } from '../projection/mdx';
import {
  parseTargetMarkdown,
  type ParsedTargetMarkdown,
  type TargetAstNode,
} from '../projection/parse-target';
import type { ReviewViewOptions } from './document-view';
import { projectPendingProperties } from './property-projection';
import { readStoredPropertyProposals } from './stored-property-proposals';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/** Materialize supported proposed content in a detached read; no live proposal is decided. */
function workingReadPackage(pkg: OoxmlPackage, revision: RevisionToken) {
  const stored = readStoredPropertyProposals(pkg, revision);
  if (!stored.success) return stored;
  const projected = projectPendingProperties(pkg, revision, stored.pending, 'working');
  if (!projected.success) return projected;
  const part = projected.pkg.parts.get(pkg.mainDocumentPart)!;
  const siteNodeIds = collectRevisionSites(part)
    .filter((site) => !site.refused)
    .map((site) => site.node.id);
  if (!siteNodeIds.length) return projected;
  const accepted = applyTreeOp(part, { op: 'acceptAllRevisions', siteNodeIds });
  return accepted.ok
    ? { success: true as const, pkg: withPart(projected.pkg, accepted.part) }
    : fail(`Unable to read proposed table geometry: ${accepted.reason}.`);
}
async function readTargetPackage(
  pkg: OoxmlPackage,
  revision: RevisionToken,
  options: ReviewViewOptions,
  syntax: 'auto' | 'gfm' = 'auto'
) {
  const store = new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!);
  const source = { ...headlessViewOfStore(store), packageRevision: () => revision.revision };
  const { pendingProperties: _pending, tableSyntax: _syntax, ...nativeOptions } = options;
  const opened = openDocumentForExport(source, {
    ...nativeOptions,
    displayMode: 'proposed',
    reuseAcrossRevisions: false,
  });
  if (!opened.ok) return fail(`Unable to read target geometry: ${opened.reason}.`);
  try {
    const layout = await opened.session.layout();
    return projectMdx(pkg, layout, revision, undefined, { tableSyntax: syntax });
  } finally {
    opened.session.dispose();
  }
}
function logicalAt(
  blocks: readonly LogicalBlock[],
  path: readonly number[]
): LogicalBlock | undefined {
  let result: LogicalBlock | undefined;
  for (const index of path) {
    result = blocks[index];
    blocks = result && 'children' in result ? result.children : [];
  }
  return result;
}
function astAt(root: TargetAstNode, path: readonly number[]): TargetAstNode | undefined {
  let result: TargetAstNode | undefined = root;
  for (const index of path)
    result = (result?.children as readonly TargetAstNode[] | undefined)?.[index];
  return result;
}

/**
 * Interpret implicit GFM geometry against the current working native table before owned inverses.
 * Only retained GFM table source ranges change. Other target bytes and explicit HTML stay exact.
 * Native preview packages supply read data only; the replacement stage still executes real operations.
 */
export async function normalizeRecompositionTarget(
  pkg: OoxmlPackage,
  revision: RevisionToken,
  target: ParsedTargetMarkdown,
  view: ReviewViewOptions,
  options: TargetDocumentOptions | undefined,
  fresh: () => NativeEditFailure | null
): Promise<
  | NativeEditFailure
  | { readonly success: true; readonly markdown: string; readonly unchanged: boolean }
> {
  const logical = normalizeLogicalBlocks(target, options?.definitions);
  if (!logical.success) return logical;
  const tables: LogicalBlock[] = [];
  const walk = (blocks: readonly LogicalBlock[]) => {
    for (const block of blocks) {
      if (block.kind === 'table' && block.header.type === 'table') tables.push(block);
      if ('children' in block) walk(block.children);
    }
  };
  walk(logical.document.blocks);
  if (!tables.length) return { success: true, markdown: target.markdown, unchanged: false };
  const visual = workingReadPackage(pkg, revision);
  if (!visual.success) return visual;
  const read = await readTargetPackage(visual.pkg, revision, view);
  if (!read.success) return read;
  const changed = fresh();
  if (changed) return changed;
  let planningMarkdown = target.markdown;
  if (read.projection.cells.some((cell) => cell.paragraphIds.length > 1)) {
    const gfm = await readTargetPackage(visual.pkg, revision, view, 'gfm');
    if (!gfm.success) return gfm;
    const changedDuringSyntax = fresh();
    if (changedDuringSyntax) return changedDuringSyntax;
    planningMarkdown = preserveUnchangedTableSyntax(target, gfm.projection, read.projection);
  }
  const plan = planTargetDocument(
    { pkg: visual.pkg, revision, projection: read.projection },
    planningMarkdown,
    options
  );
  if (!plan.success) return plan;
  const changedByPlanner = fresh();
  if (changedByPlanner) return changedByPlanner;
  if (!plan.plan.operations.length)
    return { success: true, markdown: target.markdown, unchanged: true };
  const preview = previewDirectPlan(visual.pkg, plan.plan);
  if (!preview.success) return preview;
  const proposed = await readTargetPackage(preview.pkg, revision, view);
  if (!proposed.success) return proposed;
  const changedDuringPreview = fresh();
  if (changedDuringPreview) return changedDuringPreview;
  const replacements: { start: number; end: number; text: string }[] = [];
  const ranges = new Map(
    proposed.projection.sourceMap.tables
      .filter((table) => table.sourceScope === 'body')
      .map((table) => [table.tableId, table])
  );
  for (const pair of plan.pairs) {
    const table = logicalAt(logical.document.blocks, pair.afterPath);
    if (!table || table.kind !== 'table' || table.header.type !== 'table') continue;
    const range = ranges.get(pair.id);
    if (!range) return fail('A retained target table has no complete preview source range.');
    const text = proposed.projection.text.slice(range.markdownStart, range.markdownEnd);
    if (!text.trimStart().startsWith('<table')) continue;
    const node = astAt(target.ast, table.path),
      position = node?.position as
        | { start?: { offset?: number }; end?: { offset?: number } }
        | undefined;
    const start = position?.start?.offset,
      end = position?.end?.offset;
    if (start === undefined || end === undefined)
      return fail('A retained GFM table has no complete target range.');
    replacements.push({ start, end, text });
  }
  let markdown = target.markdown,
    next = markdown.length;
  for (const item of replacements.sort((a, b) => b.start - a.start)) {
    if (item.end > next) return fail('Retained target table ranges overlap.');
    markdown = markdown.slice(0, item.start) + item.text + markdown.slice(item.end);
    next = item.start;
  }
  const checked = parseTargetMarkdown(markdown);
  return checked.success ? { success: true, markdown, unchanged: false } : checked;
}
