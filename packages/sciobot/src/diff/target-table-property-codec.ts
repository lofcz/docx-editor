import type { NativeEditFailure } from '../contracts/document';
import type { JsonObject } from '../data/json';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import { targetTableHeaderKey } from './target-table-fragments';
import {
  staticPropertyValue,
  targetLength,
  targetColor,
  targetHorizontal,
  targetVertical,
  targetBorder,
  TargetPropertyError,
  type HorizontalTarget,
  type BorderTarget,
  type NativeTargetColor,
} from './target-property-values';

export interface TargetTablePropertyValues {
  width?: number;
  height?: number;
  horizontal?: HorizontalTarget;
  vertical?: 'top' | 'center' | 'bottom';
  fill?: NativeTargetColor | null;
  borders?: Partial<Record<'outside' | 'top' | 'bottom' | 'left' | 'right', BorderTarget>>;
  columns?: readonly (HorizontalTarget | undefined)[];
}
interface Decoded {
  readonly values: TargetTablePropertyValues;
  readonly remainder: JsonObject;
}
export interface TargetTablePropertyCodec {
  readonly values: (block: LogicalGroup) => TargetTablePropertyValues;
  readonly key: (block: LogicalGroup) => string;
  readonly contentKey: (block: LogicalBlock) => string;
}

function decode(block: LogicalGroup): Decoded {
  const remainder = { ...block.header },
    values: TargetTablePropertyValues = {};
  if (!['table', 'row', 'cell'].includes(block.kind)) return { values, remainder };
  const assign = (name: string, value: string, css = false): boolean => {
    if (name === 'width' && (block.kind === 'table' || block.kind === 'cell'))
      values.width = targetLength(value, true, !css);
    else if (name === 'height' && block.kind === 'row')
      values.height = targetLength(value, true, !css);
    else if (name === 'text-align' || (name === 'align' && block.kind !== 'table'))
      values.horizontal = targetHorizontal(value);
    else if (block.kind === 'cell' && (name === 'vertical-align' || name === 'valign'))
      values.vertical = targetVertical(value);
    else if (block.kind === 'cell' && (name === 'background-color' || name === 'bgcolor'))
      values.fill = targetColor(value);
    else if (
      (block.kind === 'cell' || block.kind === 'table') &&
      /^border(?:-(top|bottom|left|right))?$/.test(name)
    ) {
      const side =
          name === 'border'
            ? 'outside'
            : (name.slice('border-'.length) as 'top' | 'bottom' | 'left' | 'right'),
        target = targetBorder(value);
      values.borders =
        side === 'outside' ? { outside: target } : { ...values.borders, [side]: target };
    } else return false;
    return true;
  };
  if (block.kind === 'table' && Array.isArray(remainder.align)) {
    if (
      remainder.align.length !== (block.children[0] as LogicalGroup | undefined)?.children.length ||
      remainder.align.some(
        (value) => value !== null && !['left', 'center', 'right'].includes(String(value))
      )
    )
      throw new TargetPropertyError('Invalid GFM column alignment metadata.');
    values.columns = remainder.align.map((value) =>
      value === null ? undefined : targetHorizontal(String(value))
    );
    delete remainder.align;
  }
  if (Array.isArray(remainder.attributes)) {
    const kept: JsonObject[string][] = [],
      styles: string[] = [],
      seen = new Set<string>();
    for (const raw of remainder.attributes) {
      const attribute = raw as JsonObject;
      if (
        !attribute ||
        attribute.type !== 'mdxJsxAttribute' ||
        typeof attribute.name !== 'string'
      ) {
        kept.push(raw);
        continue;
      }
      const name = attribute.name.toLowerCase();
      if (name === 'style') {
        if (seen.has(name)) throw new TargetPropertyError('Duplicate table property attributes.');
        seen.add(name);
        styles.push(staticPropertyValue(attribute.value));
        continue;
      }
      if (['width', 'height', 'align', 'valign', 'bgcolor'].includes(name)) {
        if (seen.has(name)) throw new TargetPropertyError('Duplicate table property attributes.');
        seen.add(name);
        const value = staticPropertyValue(attribute.value);
        if (assign(name, value)) continue;
      }
      kept.push(raw);
    }
    for (const style of styles) {
      if (style.length > 16384) throw new TargetPropertyError('Table styles exceed their limit.');
      const unknown: string[] = [],
        declarations = style.split(';').filter((value) => value.trim()),
        names = new Set<string>();
      if (declarations.length > 64)
        throw new TargetPropertyError('Table styles exceed their declaration limit.');
      for (const declaration of declarations) {
        const colon = declaration.indexOf(':');
        if (colon < 1) throw new TargetPropertyError('Malformed table style declaration.');
        const name = declaration.slice(0, colon).trim().toLowerCase(),
          value = declaration.slice(colon + 1).trim();
        if (names.has(name) || !value)
          throw new TargetPropertyError('Duplicate or empty table style declaration.');
        names.add(name);
        if (!assign(name, value, true)) unknown.push(`${name}: ${value}`);
      }
      if (unknown.length)
        kept.push({ type: 'mdxJsxAttribute', name: 'style', value: unknown.join('; ') });
    }
    remainder.attributes = kept;
  }
  return { values, remainder };
}

/** Separate owned properties from structural metadata before any content plan can return. */
export function prepareTargetTableProperties(
  before: LogicalGroup,
  after: LogicalGroup
): { readonly success: true; readonly codec: TargetTablePropertyCodec } | NativeEditFailure {
  try {
    const decoded = new Map<LogicalGroup, Decoded>();
    const visit = (block: LogicalBlock) => {
      if (!('children' in block)) return;
      decoded.set(block, decode(block));
      block.children.forEach(visit);
    };
    visit(before);
    visit(after);
    const key = (block: LogicalGroup) =>
        targetTableHeaderKey({ ...block, header: decoded.get(block)!.remainder }),
      contentKey = (block: LogicalBlock): string =>
        'children' in block
          ? JSON.stringify([key(block), block.children.map(contentKey)])
          : block.key;
    return {
      success: true,
      codec: { values: (block) => decoded.get(block)!.values, key, contentKey },
    };
  } catch (error) {
    return {
      success: false,
      code: 'unsupported-operation',
      error: error instanceof Error ? error.message : 'Invalid table property metadata.',
    };
  }
}
