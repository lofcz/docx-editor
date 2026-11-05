import type { JsonObject } from '../data/json';

/** Version 2 gives native paragraphs sole ownership of code source. */
export function codeRecordShape(value: JsonObject): boolean {
  const attributes = value.attributes;
  if (!attributes || typeof attributes !== 'object' || Array.isArray(attributes)) return false;
  const record = attributes as JsonObject;
  return (
    value.type === 'code_block' &&
    typeof record.language === 'string' &&
    record.language.length <= 128 &&
    typeof record.value === 'string'
  );
}
