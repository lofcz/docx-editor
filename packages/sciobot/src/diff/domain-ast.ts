import type { JsonObject } from '../data/json';
import type { TargetAstNode } from '../projection/parse-target';
import type { DomainNode } from './domain-strategies';
import type { MdxComponentRegistry } from '../registry/contract';
import { readMdxComponent } from '../registry/mdx-read';
import { activeMdxRegistry } from '../registry/active';

/**
 * The payload node a registered JSX element describes, or null for any other element.
 *
 * Interpret data attributes only. MDX expressions remain inert JSON objects, never JavaScript
 * values. Each component's parser owns its attribute names and normalization.
 */
export function domainNodeFromAst(
  ast: TargetAstNode,
  registry: MdxComponentRegistry = activeMdxRegistry()
): (DomainNode & JsonObject) | null {
  return readMdxComponent(registry, ast);
}
