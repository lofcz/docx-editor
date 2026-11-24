import { contentControlsIn, type OoxmlElement, type OoxmlNode } from '@docx-editor.dev/core/store';
import type {
  InlineWidgetMetrics,
  InlineWidgetMetricsProvider,
} from '../contracts/inline-widget-metrics.ts';
import { framedTokenJoin } from './framed-token.ts';
import { inlineWidgetProducer } from './inline-widget-layout.ts';
import type { SemanticLayoutOptions } from './semantic-layout-options.ts';
import { styleSeparatorSourceBlocks } from './style-separator-group.ts';
import {
  aggregateParagraphTokensForTableBlock,
  tableParagraphsInOrder,
} from './table-paragraph-tokens.ts';

const byProvider = new WeakMap<
  InlineWidgetMetricsProvider,
  ReadonlyMap<string, InlineWidgetMetrics>
>();

/** Capture a complete host map. Neither later map mutations nor box mutations change a layout pass. */
export function captureInlineWidgetMetrics(boxes: ReadonlyMap<string, InlineWidgetMetrics>) {
  const copy = new Map<string, InlineWidgetMetrics>();
  for (const [id, box] of boxes)
    copy.set(id, Object.freeze({ width: box.width, ascent: box.ascent, descent: box.descent }));
  const measure: InlineWidgetMetricsProvider = ({ controlId }) => copy.get(controlId) ?? null;
  byProvider.set(measure, copy);
  return { measure, revision: '' };
}

const owners = new WeakMap<readonly OoxmlElement[], ReadonlyMap<string, string>>();
function directOwners(blocks: readonly OoxmlElement[]) {
  const saved = owners.get(blocks);
  if (saved) return saved;
  const found = new Map<string, string>();
  const record = (paragraph: OoxmlNode) => {
    if (paragraph.kind !== 'paragraph') return;
    for (const child of paragraph.children) {
      if (child.kind === 'contentControl') found.set(child.id, paragraph.id);
    }
  };
  for (const block of styleSeparatorSourceBlocks(blocks)) {
    // Cell paragraphs own their inline w:sdt (CT_SdtRun, ECMA-376 §17.5.2.31).
    // Reuse the table walk that keys w:tc (CT_Tc, §17.4.65) for layout.
    if (block.kind === 'table') {
      for (const paragraph of tableParagraphsInOrder(block)) record(paragraph);
    } else record(block);
  }
  owners.set(blocks, found);
  return found;
}

const projections = new WeakMap<
  ReadonlyMap<string, InlineWidgetMetrics>,
  WeakMap<readonly OoxmlElement[], ReturnType<typeof tokensForBlocks>>
>();
function tokensForBlocks(
  boxes: ReadonlyMap<string, InlineWidgetMetrics>,
  blocks: readonly OoxmlElement[]
) {
  const paragraphs = directOwners(blocks);
  const entries = new Map<string, string[]>();
  const global: string[] = [];
  for (const [id, box] of boxes) {
    const owner = paragraphs.get(id);
    const key = JSON.stringify([id, box.width, box.ascent, box.descent]);
    if (owner === undefined) {
      global.push(key);
      continue;
    }
    const tokens = entries.get(owner) ?? [];
    tokens.push(key);
    entries.set(owner, tokens);
  }
  return {
    global: framedTokenJoin(global.sort()),
    paragraphs: new Map([...entries].map(([id, keys]) => [id, framedTokenJoin(keys.sort())])),
    tables: new WeakMap<OoxmlNode, string>(),
  };
}

function hasFurnitureWidgets(furniture: SemanticLayoutOptions['furniture']) {
  if (!furniture) return false;
  for (const variants of [furniture.headers, furniture.footers]) {
    for (const story of variants.values()) {
      if (!story.part || contentControlsIn(story.part.root).length) return true;
    }
  }
  return false;
}

/** Unknown owners invalidate every block. A known box keys its paragraph and enclosing tables. */
export function indexedWidgetProjection(
  options: SemanticLayoutOptions,
  blocks: readonly OoxmlElement[]
) {
  const source = options.measurer.inlineWidgets;
  const boxes = source && byProvider.get(source.measure);
  if (
    !boxes ||
    options.notes ||
    hasFurnitureWidgets(options.furniture) ||
    options.sectionFurniture?.some(hasFurnitureWidgets)
  )
    return undefined;
  let cached = projections.get(boxes);
  if (!cached) projections.set(boxes, (cached = new WeakMap()));
  let tokens = cached.get(blocks);
  if (tokens === undefined) {
    tokens = tokensForBlocks(boxes, blocks);
    cached.set(blocks, tokens);
  }
  const previous = options.projectionTokenForParagraph;
  const widgetToken = (node: OoxmlNode) => tokens.paragraphs.get(node.id) ?? '';
  return {
    global: tokens.global,
    inputs: {
      projectionEpoch:
        previous && options.projectionEpoch === undefined
          ? undefined
          : framedTokenJoin([
              options.projectionEpoch ?? '',
              inlineWidgetProducer('', options.measurer)!,
            ]),
      projectionTokenForParagraph: (p: OoxmlNode) => {
        const prior = previous?.(p) ?? '';
        const token = widgetToken(p);
        return token ? framedTokenJoin([prior, token]) : prior;
      },
      projectionTokenForTable: (table: OoxmlNode) => {
        const prior =
          options.projectionTokenForTable?.(table) ??
          (previous ? aggregateParagraphTokensForTableBlock(table, previous) : '');
        let token = tokens.tables.get(table);
        if (token === undefined) {
          token = aggregateParagraphTokensForTableBlock(table, widgetToken);
          tokens.tables.set(table, token);
        }
        return token ? framedTokenJoin([prior, token]) : prior;
      },
    },
  };
}
