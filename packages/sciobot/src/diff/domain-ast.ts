import type { JsonObject, JsonValue } from '../data/json';
import type { TargetAstNode } from '../projection/parse-target';
import { IGNORED_DIFF_PROPERTIES, type DomainNode } from './domain-strategies';

const tags: Readonly<Record<string, string>> = {
  phase: 'lesson_phase',
  activity: 'lesson_activity',
  lesson_title: 'lesson_title',
  preparation_recap: 'preparation_recap',
  recap_goals: 'recap_goals',
  recap_curriculum: 'recap_curriculum',
  recap_materials: 'recap_materials',
  recap_adaptations: 'recap_adaptations',
  recap_phase: 'recap_phase',
  lesson_visual: 'lesson_visual',
  listening: 'listening',
};
function hasText(node: TargetAstNode): boolean {
  return (
    (typeof node.value === 'string' && node.value.trim().length > 0) ||
    (Array.isArray(node.children) && (node.children as TargetAstNode[]).some(hasText))
  );
}
/** Interpret data attributes only. MDX expressions remain inert JSON objects, never JavaScript values. */
export function domainNodeFromAst(ast: TargetAstNode): (DomainNode & JsonObject) | null {
  if (
    !['mdxJsxFlowElement', 'mdxJsxTextElement'].includes(ast.type) ||
    typeof ast.name !== 'string'
  )
    return null;
  const type =
    ast.name === 'source'
      ? hasText(ast)
        ? 'source_excerpt'
        : 'source'
      : Object.hasOwn(tags, ast.name)
        ? tags[ast.name]
        : undefined;
  if (!type) return null;
  const attributes: [string, JsonValue][] = [];
  for (const raw of Array.isArray(ast.attributes) ? ast.attributes : []) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) continue;
    const attr = raw as JsonObject;
    if (attr.type !== 'mdxJsxAttribute' || typeof attr.name !== 'string') continue;
    let name = attr.name,
      value = attr.value;
    if (name === 'type' && (type === 'lesson_phase' || type === 'recap_phase')) name = 'phaseType';
    else if (name === 'type' && (type === 'source' || type === 'source_excerpt')) name = 'trust';
    else if (name === 'organizational_minutes' && type === 'lesson_title')
      name = 'organizationalMinutes';
    else if (name === 'lesson_minutes' && type === 'lesson_title') name = 'lessonMinutes';
    if (
      name === 'type' ||
      name === 'children' ||
      IGNORED_DIFF_PROPERTIES.has(name) ||
      value === undefined
    )
      continue;
    if (value === null) value = true;
    else if (typeof value === 'string') {
      try {
        value = JSON.parse(value) as JsonValue;
      } catch {
        /* Literal text. */
      }
    }
    attributes.push([name, value]);
  }
  const props = Object.fromEntries(attributes) as Record<string, JsonValue>;
  if (type === 'source' || type === 'source_excerpt') {
    props.url = typeof props.url === 'string' ? props.url : '';
    if (typeof props.title !== 'string') delete props.title;
    if (typeof props.author !== 'string') delete props.author;
    const book = (value: JsonValue | undefined): number | null =>
      typeof value === 'number' && Number.isFinite(value) && value > 0
        ? Math.trunc(value)
        : typeof value === 'string' && /^\d{1,12}$/.test(value.trim()) && Number(value) > 0
          ? Number(value)
          : null;
    const bookId = book(props.bookId) ?? book(props.bookid);
    delete props.bookid;
    if (bookId === null) delete props.bookId;
    else props.bookId = bookId;
    if (props.trust === 'web' || props.trust === 'youtube') {
      if (props.url) props.trust = 'verified';
      else delete props.trust;
    } else if (props.trust !== 'verified' && props.trust !== 'generated') delete props.trust;
  }
  return { ...props, type };
}
