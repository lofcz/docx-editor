import type { MdxComponentRegistry } from '../registry/contract';
import { sciobotMdxRegistry } from '../registry/sciobot';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { MDX_BLOCK } from '../projection/mdx';
import type { TargetAstNode } from '../projection/parse-target';
import { compileInsertedAtoms, type InsertedInlineAtom } from './inserted-atoms';
import type { InsertedBlockBinding } from './inserted-blocks';
import { insertedPayloadProblem } from './inserted-payloads';
import { jsonDataKey } from './target-key';

/** Cell paragraphs use one final paragraph index after the row/cell path, including implicit GFM paragraphs. */
export function insertedParagraphs(
  inserted: readonly TargetAstNode[],
  blocks: readonly InsertedBlockBinding[],
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): readonly { readonly path: readonly number[]; readonly ast: TargetAstNode }[] {
  const atomic = new Set(
    blocks
      .filter(
        (binding) =>
          binding.payload.type === MDX_BLOCK ||
          registry.blockTags.get(binding.payload.type)?.kind === 'atomic'
      )
      .map((binding) => binding.path.join('/'))
  );
  const result: { path: readonly number[]; ast: TargetAstNode }[] = [];
  const visit = (node: TargetAstNode, path: readonly number[]) => {
    if (atomic.has(path.join('/'))) return;
    if (node.type === 'paragraph' || node.type === 'heading') {
      result.push({ path, ast: node });
      return;
    }
    if (node.type === 'tableCell') {
      const children = (node.children ?? []) as TargetAstNode[];
      const paragraphs = children.every((child) => child.type === 'paragraph')
        ? children
        : [{ type: 'paragraph', children } as TargetAstNode];
      paragraphs.forEach((child, index) => visit(child, [...path, index]));
      return;
    }
    if (Array.isArray(node.children))
      (node.children as TargetAstNode[]).forEach((child, index) => visit(child, [...path, index]));
  };
  inserted.forEach((node, index) => visit(node, [index]));
  return result;
}
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'invalid-input',
  error,
});

export function planInsertedAtomBindings(
  pkg: OoxmlPackage,
  inserted: readonly TargetAstNode[],
  blocks: readonly InsertedBlockBinding[],
  allocate: () => string = () => crypto.randomUUID(),
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): { readonly success: true; readonly atoms: readonly InsertedInlineAtom[] } | NativeEditFailure {
  const atoms: InsertedInlineAtom[] = [];
  const used = new Set(blocks.map((binding) => binding.payload.id));
  for (const paragraph of insertedParagraphs(inserted, blocks, registry)) {
    const compiled = compileInsertedAtoms(
      {
        location: { kind: 'new-paragraph', path: paragraph.path },
        offset: 0,
        allocateBlockId: allocate,
      },
      paragraph.ast
    );
    if (!compiled.success) return compiled;
    for (const op of compiled.plan.operations) {
      if (op.kind === 'inline-control') {
        if (used.has(op.payload.id)) return fail('Inserted block and atom IDs must be distinct.');
        used.add(op.payload.id);
      }
      atoms.push(op);
    }
  }
  const controls = atoms.flatMap((op) =>
    op.kind === 'inline-control' ? [{ path: op.path, payload: op.payload, label: op.label }] : []
  );
  const problem = insertedPayloadProblem(pkg, controls);
  return problem ?? { success: true, atoms };
}

/** The D01 decoder reconstructs intentions from AST and declared IDs; no package writes occur. */
export function isInsertedAtomBindings(
  inserted: readonly TargetAstNode[],
  blocks: readonly InsertedBlockBinding[],
  value: unknown,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): boolean {
  if (!Array.isArray(value)) return false;
  try {
    const controls = (value as InsertedInlineAtom[]).filter((op) => op.kind === 'inline-control');
    const ids = [
      ...blocks.map((binding) => binding.payload.id),
      ...controls.map((op) => op.payload.id),
    ];
    if (new Set(ids).size !== ids.length) return false;
    let next = 0;
    const expected: InsertedInlineAtom[] = [];
    for (const paragraph of insertedParagraphs(inserted, blocks, registry)) {
      const compiled = compileInsertedAtoms(
        {
          location: { kind: 'new-paragraph', path: paragraph.path },
          offset: 0,
          allocateBlockId: () => controls[next++]?.payload.id ?? '',
        },
        paragraph.ast
      );
      if (!compiled.success) return false;
      expected.push(...compiled.plan.operations);
    }
    return (
      next === controls.length &&
      jsonDataKey(expected as unknown as import('../data/json').JsonValue) === jsonDataKey(value)
    );
  } catch {
    return false;
  }
}
