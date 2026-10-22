import type { JsonObject, JsonValue } from '../data/json';

/** Call after the bounded JSON gate. Object order is incidental; array order and opaque data are exact. */
function keyOf(value: JsonValue, ast: boolean): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return '[' + value.map((child) => keyOf(child, ast)).join(',') + ']';
  const node = value as JsonObject,
    isNode = ast && typeof node.type === 'string';
  return (
    '{' +
    Object.keys(node)
      .filter((key) => !isNode || key !== 'position')
      .sort()
      .map((key) => {
        const child = node[key]!;
        const nestedAst =
          isNode &&
          (((key === 'children' || key === 'attributes') && Array.isArray(child)) ||
            (key === 'value' &&
              child !== null &&
              typeof child === 'object' &&
              !Array.isArray(child) &&
              (child as JsonObject).type === 'mdxJsxAttributeValueExpression'));
        return JSON.stringify(key) + ':' + keyOf(child, nestedAst);
      })
      .join(',') +
    '}'
  );
}

export const jsonDataKey = (value: JsonValue): string => keyOf(value, false);
/** Ignore locations only along AST edges. A typed object inside `data` remains opaque data. */
export const targetAstKey = (value: JsonValue): string => keyOf(value, true);
