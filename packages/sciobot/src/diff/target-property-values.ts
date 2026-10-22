import type { TreeDocOp } from '@docx-editor.dev/core/store';

export type NativeTargetColor = Exclude<
  Extract<TreeDocOp, { op: 'setTableCellFill' }>['color'],
  null
>;
type NativeBorderSpec = Extract<
  Extract<TreeDocOp, { op: 'setTableCellBorders' }>,
  { spec: unknown }
>['spec'];

export class TargetPropertyError extends Error {}
export type HorizontalTarget = 'left' | 'center' | 'right' | 'both' | 'start' | 'end' | null;
export type BorderTarget =
  | { readonly clear: true }
  | { readonly clear: false; readonly spec: NativeBorderSpec };
const invalid = (message: string): never => {
  throw new TargetPropertyError(message);
};

/** Only literal strings/numbers are accepted. MDX expressions are parsed as JSON, never evaluated. */
export function staticPropertyValue(value: unknown): string {
  if (typeof value === 'string') return value;
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    const attribute = value as Record<string, unknown>;
    if (
      attribute.type === 'mdxJsxAttributeValueExpression' &&
      typeof attribute.value === 'string'
    ) {
      let literal: unknown;
      try {
        literal = JSON.parse(attribute.value);
      } catch {
        return invalid('Table properties require static literal values.');
      }
      if (typeof literal === 'string' || (typeof literal === 'number' && Number.isFinite(literal)))
        return String(literal);
    }
  }
  return invalid('Table properties require static string or numeric values.');
}

/** CSS pixels use 96 dpi. Native row/grid lengths round once to integer twips. */
export function targetLength(value: string, round = true, unitless = true): number {
  if (value.length > 64) return invalid('Table lengths exceed their limit.');
  const match = /^\s*(\d+(?:\.\d+)?)\s*(px|pt|in|cm|mm)?\s*$/i.exec(value);
  if (!match) return invalid('Table lengths need absolute px, pt, in, cm, or mm units.');
  if (!unitless && !match[2] && Number(match[1]) !== 0)
    return invalid('CSS table lengths need an explicit absolute unit.');
  const factors: Record<string, number> = {
      px: 15,
      pt: 20,
      in: 1440,
      cm: 1440 / 2.54,
      mm: 1440 / 25.4,
    },
    result = Number(match[1]) * factors[match[2]?.toLowerCase() ?? 'px']!;
  if (!Number.isFinite(result) || result > 31680)
    return invalid('Table lengths exceed the native limit.');
  return round ? Math.round(result) : result;
}

const named: Readonly<Record<string, string>> = {
  black: '000000',
  silver: 'C0C0C0',
  gray: '808080',
  white: 'FFFFFF',
  maroon: '800000',
  red: 'FF0000',
  purple: '800080',
  fuchsia: 'FF00FF',
  green: '008000',
  lime: '00FF00',
  olive: '808000',
  yellow: 'FFFF00',
  navy: '000080',
  blue: '0000FF',
  teal: '008080',
  aqua: '00FFFF',
  orange: 'FFA500',
  rebeccapurple: '663399',
};
export function targetColor(value: string): NativeTargetColor | null {
  const color = value.trim().toLowerCase();
  if (!color || color === 'transparent') return null;
  let hex =
    /^#?([\da-f]{6})$/i.exec(color)?.[1] ??
    (Object.hasOwn(named, color) ? named[color] : undefined);
  if (!hex) {
    const short = /^#([\da-f]{3})$/i.exec(color)?.[1];
    if (short) hex = [...short].map((value) => value + value).join('');
  }
  if (!hex) {
    const rgb = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/.exec(color);
    if (rgb && rgb.slice(1).every((part) => Number(part) <= 255))
      hex = rgb
        .slice(1)
        .map((part) => Number(part).toString(16).padStart(2, '0'))
        .join('');
  }
  if (!hex) return invalid('Table colors need an explicit supported opaque color or transparent.');
  return { kind: 'hex', value: hex.toUpperCase() };
}
export function targetHorizontal(value: string): HorizontalTarget {
  const alignment = value.trim().toLowerCase();
  if (!alignment || alignment === 'inherit' || alignment === 'unset') return null;
  if (alignment === 'justify') return 'both';
  if (['left', 'center', 'right', 'start', 'end'].includes(alignment))
    return alignment as Exclude<HorizontalTarget, 'both' | null>;
  return invalid('Unsupported native horizontal alignment.');
}
export function targetVertical(value: string): 'top' | 'center' | 'bottom' {
  const alignment = value.trim().toLowerCase();
  if (alignment === 'middle' || alignment === 'center') return 'center';
  if (alignment === 'top' || alignment === 'bottom') return alignment;
  return invalid('Unsupported native vertical alignment.');
}
export function targetBorder(value: string): BorderTarget {
  if (value.length > 256) return invalid('Table border values exceed their limit.');
  const lower = value.trim().toLowerCase();
  if (lower === 'none' || lower === '0') return { clear: true };
  const parts = lower.match(/rgb\([^)]*\)|[^\s]+/g) ?? [],
    styles = { solid: 'single', double: 'double', dotted: 'dotted', dashed: 'dashed' } as const,
    style = parts.find((part) => Object.hasOwn(styles, part)) as keyof typeof styles | undefined,
    width = parts.find((part) => /^\d/.test(part));
  if (parts.length !== 3 || !style || !width)
    return invalid('Table borders need width, supported style, and color.');
  const authoredColor = parts.find((part) => part !== style && part !== width);
  if (!authoredColor) return invalid('Table borders need an explicit color.');
  const color = targetColor(authoredColor);
  if (!color) return invalid('A drawn table border needs an opaque color.');
  const size = Math.round((targetLength(width, false, false) * 8) / 20);
  return size ? { clear: false, spec: { style: styles[style], size, color } } : { clear: true };
}
