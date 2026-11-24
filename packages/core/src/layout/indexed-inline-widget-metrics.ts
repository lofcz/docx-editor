import type { OoxmlElement, OoxmlNode } from '@docx-editor.dev/core/store';
import type { InlineWidgetMetrics, InlineWidgetMetricsProvider } from '../contracts/inline-widget-metrics.ts';
import { framedTokenJoin } from './framed-token.ts';
import { inlineWidgetProducer } from './inline-widget-layout.ts';
import type { SemanticLayoutOptions } from './semantic-layout-options.ts';
import { styleSeparatorSourceBlocks } from './style-separator-group.ts';

const byProvider = new WeakMap<InlineWidgetMetricsProvider, ReadonlyMap<string, InlineWidgetMetrics>>();

/** Capture a complete host map. Neither later map mutations nor box mutations change a layout pass. */
export function captureInlineWidgetMetrics(boxes: ReadonlyMap<string, InlineWidgetMetrics>) {
  const copy = new Map<string, InlineWidgetMetrics>();
  for (const [id, box] of boxes) copy.set(id, Object.freeze({width: box.width, ascent: box.ascent, descent: box.descent}));
  const measure: InlineWidgetMetricsProvider = ({controlId}) => copy.get(controlId) ?? null;
  byProvider.set(measure, copy);
  return {measure, revision: ''};
}

const owners = new WeakMap<readonly OoxmlElement[], ReadonlyMap<string, string>>();
function directOwners(blocks: readonly OoxmlElement[]) {
  const saved = owners.get(blocks);
  if (saved) return saved;
  const found = new Map<string, string>();
  for (const paragraph of styleSeparatorSourceBlocks(blocks)) {
    if (paragraph.kind !== 'paragraph') continue;
    for (const child of paragraph.children) {
      if (child.kind === 'contentControl') found.set(child.id, paragraph.id);
    }
  }
  owners.set(blocks, found);
  return found;
}

/** Use paragraph keys only when every published box has a known direct body owner. */
export function indexedWidgetProjection(options: SemanticLayoutOptions, blocks: readonly OoxmlElement[]) {
  const source = options.measurer.inlineWidgets;
  const boxes = source && byProvider.get(source.measure);
  const hasFurniture = (f: SemanticLayoutOptions['furniture']) => Boolean(f && (f.headers.size || f.footers.size));
  if (!boxes || options.notes || hasFurniture(options.furniture) || options.sectionFurniture?.some(hasFurniture)) return undefined;
  const paragraphs = directOwners(blocks);
  const entries = new Map<string, string[]>();
  for (const [id, box] of boxes) {
    const owner = paragraphs.get(id);
    if (owner === undefined) return undefined;
    const tokens = entries.get(owner) ?? [];
    tokens.push(JSON.stringify([id,box.width,box.ascent,box.descent]));
    entries.set(owner,tokens);
  }
  const tokens = new Map([...entries].map(([id, keys]) => [id, framedTokenJoin(keys.sort())]));
  const previous = options.projectionTokenForParagraph;
  return {
    projectionEpoch: previous && options.projectionEpoch === undefined ? undefined : framedTokenJoin([options.projectionEpoch ?? '',inlineWidgetProducer('',options.measurer)!]),
    projectionTokenForParagraph: (p: OoxmlNode) => {
      const prior = previous?.(p) ?? '';
      const token = tokens.get(p.id);
      return token ? framedTokenJoin([prior,token]) : prior;
    },
  };
}
