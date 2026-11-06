import { ATTRIBUTE_FIELD, fieldChanges } from '../blocks/payload-fields';
import type { BlockPayload, JsonObject, JsonValue } from '../blocks/payload-schema';
import { jsonDataKey } from '../diff/target-key';

export const nativeCodePayload = (payload: BlockPayload): boolean =>
  payload.type === 'code_block' && payload.version === 2;

const valueKey = (object: JsonObject, name: string): string =>
  Object.hasOwn(object, name) ? `present:${jsonDataKey(object[name]!)}` : 'absent';

/** Native sites prove source ownership. Only declared metadata changes belong to the payload proof. */
export function codeMetadataProof(
  current: BlockPayload,
  before: BlockPayload,
  after: BlockPayload
) {
  if (
    ![current, before, after].every(nativeCodePayload) ||
    before.id !== current.id ||
    after.id !== current.id
  )
    throw Error('Native code evidence requires matching version-2 payload identities.');
  return fieldChanges(before, after)
    .filter((field) => field.kind !== ATTRIBUTE_FIELD || field.name !== 'value')
    .map((field) => [
      field.kind,
      field.name,
      valueKey(field.kind === ATTRIBUTE_FIELD ? current.attributes : current, field.name),
    ]);
}

/** Native code source is never a payload override. Unchanged metadata retains its current owner. */
export function projectedCodeMetadata(
  current: BlockPayload,
  states: readonly BlockPayload[],
  view: 'committed' | 'working'
): BlockPayload {
  if (
    !nativeCodePayload(current) ||
    states.length < 2 ||
    states.some((state) => !nativeCodePayload(state) || state.id !== current.id)
  )
    throw Error('Native code review requires matching version-2 payload identities.');
  const fields = new Map<string, ReturnType<typeof fieldChanges>[number]>();
  for (let index = 1; index < states.length; index++)
    for (const field of fieldChanges(states[index - 1]!, states[index]!)) {
      if (field.kind === ATTRIBUTE_FIELD && field.name === 'value') continue;
      fields.set(`${field.kind}:${field.name}`, field);
    }
  if (!fields.size) return current;
  const selected = view === 'working' ? states.at(-1)! : states[0]!;
  const attributes = { ...current.attributes };
  const next: Record<string, JsonValue> = { ...current, attributes };
  for (const field of fields.values()) {
    const objectOf = (payload: BlockPayload) =>
      field.kind === ATTRIBUTE_FIELD ? payload.attributes : payload;
    const actual = valueKey(objectOf(current), field.name);
    if (!states.some((state) => valueKey(objectOf(state), field.name) === actual))
      throw Error('A proposed code metadata field changed outside its retained history.');
    const source = objectOf(selected);
    const target = field.kind === ATTRIBUTE_FIELD ? attributes : next;
    if (!Object.hasOwn(source, field.name)) delete target[field.name];
    else
      Object.defineProperty(target, field.name, {
        value: source[field.name] as JsonValue,
        enumerable: true,
        configurable: true,
        writable: true,
      });
  }
  return next as BlockPayload;
}
