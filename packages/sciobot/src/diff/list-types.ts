export const LIST_STYLES = [
  'disc',
  'circle',
  'square',
  'decimal',
  'lower-alpha',
  'upper-alpha',
  'lower-roman',
  'upper-roman',
] as const;
export type ListStyle = (typeof LIST_STYLES)[number];
/** Semantic list ownership. New instances receive native IDs inside the eventual transaction. */
export type ListInstance =
  | { readonly kind: 'existing'; readonly numId: string }
  /** One key joins a multilevel instance. Kind and start belong to each target level. */
  | { readonly kind: 'new'; readonly key: string; readonly start?: number };

export type ListTarget =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'bullet' | 'ordered';
      readonly level: number;
      readonly instance: ListInstance;
      /** Style for all levels of a new counter instance; omitted uses the Word template. */
      readonly style?: ListStyle;
    }
  | { readonly kind: 'checklist'; readonly level: number; readonly checked: boolean };

export interface ListSourceState {
  readonly numbering: { readonly numId: string; readonly level: number } | null;
  readonly checkbox: {
    readonly controlId: string;
    readonly checked: boolean;
    readonly level: number;
  } | null;
  /** Authored numbering, list indentation, and owned checkbox content. Never a document tree. */
  readonly fingerprint: string;
}

const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const keys = (value: object, allowed: readonly string[]) =>
  Object.keys(value).every((key) => allowed.includes(key));
const level = (value: unknown): boolean =>
  Number.isInteger(value) && Number(value) >= 0 && Number(value) <= 8;
const identifier = (value: unknown): boolean =>
  typeof value === 'string' && value.length > 0 && value.length <= 256;

/** Called only after the D01 plain JSON gate. */
export function isListTarget(value: unknown): value is ListTarget {
  if (!record(value)) return false;
  if (value.kind === 'none') return keys(value, ['kind']);
  if (!level(value.level)) return false;
  if (value.kind === 'checklist')
    return typeof value.checked === 'boolean' && keys(value, ['kind', 'level', 'checked']);
  if (
    !['bullet', 'ordered'].includes(String(value.kind)) ||
    !keys(value, ['kind', 'level', 'instance', 'style'])
  )
    return false;
  if (
    value.style !== undefined &&
    (!LIST_STYLES.includes(value.style as ListStyle) ||
      (value.kind === 'bullet') !== ['disc', 'circle', 'square'].includes(String(value.style)))
  )
    return false;
  const instance = value.instance;
  if (!record(instance)) return false;
  if (instance.kind === 'existing')
    return (
      value.style === undefined &&
      typeof instance.numId === 'string' &&
      /^\d{1,9}$/.test(instance.numId) &&
      Number(instance.numId) > 0 &&
      keys(instance, ['kind', 'numId'])
    );
  return (
    instance.kind === 'new' &&
    identifier(instance.key) &&
    (instance.start === undefined ||
      (value.kind === 'ordered' &&
        Number.isSafeInteger(instance.start) &&
        Number(instance.start) >= 0 &&
        Number(instance.start) <= 2147483647)) &&
    keys(instance, ['kind', 'key', 'start'])
  );
}

export function isListSourceState(value: unknown): value is ListSourceState {
  if (
    !record(value) ||
    !keys(value, ['numbering', 'checkbox', 'fingerprint']) ||
    typeof value.fingerprint !== 'string' ||
    !/^sha256:[a-f0-9]{64}$/.test(value.fingerprint)
  )
    return false;
  const numbering = value.numbering,
    checkbox = value.checkbox;
  return (
    (numbering === null ||
      (record(numbering) &&
        typeof numbering.numId === 'string' &&
        /^\d{1,9}$/.test(numbering.numId) &&
        Number(numbering.numId) > 0 &&
        level(numbering.level) &&
        keys(numbering, ['numId', 'level']))) &&
    (checkbox === null ||
      (record(checkbox) &&
        typeof checkbox.controlId === 'string' &&
        checkbox.controlId.length > 0 &&
        checkbox.controlId.length <= 4096 &&
        typeof checkbox.checked === 'boolean' &&
        level(checkbox.level) &&
        keys(checkbox, ['controlId', 'checked', 'level'])))
  );
}
