import type { JsonObject, JsonValue } from '../data/json';
import type { TargetAstNode } from '../projection/parse-target';
import { IGNORED_DIFF_PROPERTIES, type DomainNode } from '../diff/domain-strategies';
import type { MdxAttribute, MdxComponentDefinition, MdxComponentRegistry } from './contract';

/** The plain JSX attributes of an element. Spread expressions and malformed entries drop out. */
export function mdxAttributes(ast: TargetAstNode): MdxAttribute[] {
  const out: MdxAttribute[] = [];
  for (const raw of Array.isArray(ast.attributes) ? ast.attributes : []) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) continue;
    const attr = raw as JsonObject;
    if (attr.type !== 'mdxJsxAttribute' || typeof attr.name !== 'string') continue;
    out.push({ name: attr.name, value: attr.value });
  }
  return out;
}

export interface MdxPropsOptions {
  /** Attribute names that become other payload names, applied before any other rule. */
  readonly rename?: Readonly<Record<string, string>>;
  /** Payload names whose string values stay literal instead of being read as JSON. */
  readonly literal?: readonly string[];
}

/**
 * Read attributes as data only. A bare attribute is `true`; a string that parses as JSON is that
 * value; anything else stays literal text. `type`, `children`, and ignored diff names drop out.
 */
export function mdxProps(
  attributes: readonly MdxAttribute[],
  options: MdxPropsOptions = {}
): Record<string, JsonValue> {
  const entries: [string, JsonValue][] = [];
  for (const attribute of attributes) {
    const name =
      options.rename && Object.hasOwn(options.rename, attribute.name)
        ? options.rename[attribute.name]!
        : attribute.name;
    let value = attribute.value;
    if (
      name === 'type' ||
      name === 'children' ||
      IGNORED_DIFF_PROPERTIES.has(name) ||
      value === undefined
    )
      continue;
    if (value === null) value = true;
    else if (typeof value === 'string' && !options.literal?.includes(name)) {
      try {
        value = JSON.parse(value) as JsonValue;
      } catch {
        /* Literal text. */
      }
    }
    entries.push([name, value]);
  }
  return Object.fromEntries(entries) as Record<string, JsonValue>;
}

/** Whether an element holds visible text, which makes a shared tag read as its container. */
export function mdxHasText(node: TargetAstNode): boolean {
  return (
    (typeof node.value === 'string' && node.value.trim().length > 0) ||
    (Array.isArray(node.children) && (node.children as TargetAstNode[]).some(mdxHasText))
  );
}

/** The registered component a JSX element names, or null for plain and unknown elements. */
export function componentForAst(
  registry: MdxComponentRegistry,
  ast: TargetAstNode
): MdxComponentDefinition | null {
  if (
    !['mdxJsxFlowElement', 'mdxJsxTextElement'].includes(ast.type) ||
    typeof ast.name !== 'string' ||
    !registry.parsers.has(ast.name)
  )
    return null;
  const entry = registry.parsers.get(ast.name)!;
  if (entry.atomic && entry.container) return mdxHasText(ast) ? entry.container : entry.atomic;
  return entry.container ?? entry.atomic ?? null;
}

/** Interpret data attributes only. MDX expressions remain inert JSON objects, never JavaScript values. */
export function readMdxComponent(
  registry: MdxComponentRegistry,
  ast: TargetAstNode
): (DomainNode & JsonObject) | null {
  const component = componentForAst(registry, ast);
  if (!component?.parse) return null;
  const props = component.parse.read(mdxAttributes(ast), ast);
  return props ? { ...props, type: component.type } : null;
}
