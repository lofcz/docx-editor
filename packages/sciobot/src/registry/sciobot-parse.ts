import type { JsonObject, JsonValue } from '../data/json';
import type { MdxAttribute, MdxComponentParser } from './contract';
import { mdxProps } from './mdx-read';

const plain = (tag: string): MdxComponentParser => ({
  tag,
  read: (attributes) => mdxProps(attributes) as JsonObject,
});

/** Phases write their payload `phaseType` as the `type` attribute. */
const phase = (tag: string): MdxComponentParser => ({
  tag,
  read: (attributes) => mdxProps(attributes, { rename: { type: 'phaseType' } }) as JsonObject,
});

/** `type` and `variant` both name the callout variant; `type` wins, both compare lowercase. */
const callout: MdxComponentParser = {
  tag: 'callout',
  read(attributes) {
    let type: string | undefined;
    let variant: string | undefined;
    const rest: MdxAttribute[] = [];
    for (const attribute of attributes) {
      if (attribute.name !== 'type' && attribute.name !== 'variant') {
        rest.push(attribute);
        continue;
      }
      const value = attribute.value;
      const normalized = typeof value === 'string' ? value.trim().toLowerCase() : '';
      if (attribute.name === 'type') type = normalized || undefined;
      else variant = normalized || undefined;
    }
    const props = mdxProps(rest);
    if (type ?? variant) props.variant = (type ?? variant)!;
    return props as JsonObject;
  },
};

function bookId(value: JsonValue | undefined): number | null {
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) return Math.trunc(value);
  if (typeof value === 'string' && /^\d{1,12}$/.test(value.trim()) && Number(value) > 0)
    return Number(value);
  return null;
}

/** The application's source grammar: trust, URL, book reference, and inferred media chrome. */
function source(tag: string): MdxComponentParser {
  return {
    tag,
    read(attributes) {
      const props = mdxProps(attributes, { rename: { type: 'trust' } });
      props.url = typeof props.url === 'string' ? props.url : '';
      if (typeof props.title !== 'string') delete props.title;
      if (typeof props.author !== 'string') delete props.author;
      const book = bookId(props.bookId) ?? bookId(props.bookid);
      delete props.bookid;
      if (book === null) delete props.bookId;
      else props.bookId = book;
      // Media chrome is inferred independently of trust.
      if (!Object.hasOwn(props, 'sourceType')) {
        let media: 'web' | 'youtube' = props.trust === 'youtube' ? 'youtube' : 'web';
        if (props.trust !== 'web' && props.trust !== 'youtube') {
          try {
            const host = new URL(props.url).hostname.replace(/^www\./, '');
            if (['youtube.com', 'm.youtube.com', 'music.youtube.com', 'youtu.be'].includes(host))
              media = 'youtube';
          } catch {
            /* Malformed URLs keep ordinary web chrome. */
          }
        }
        props.sourceType = media;
      }
      if (props.trust === 'web' || props.trust === 'youtube') {
        if (props.url) props.trust = 'verified';
        else delete props.trust;
      } else if (props.trust !== 'verified' && props.trust !== 'generated') delete props.trust;
      return props as JsonObject;
    },
  };
}

const equation = (tag: string): MdxComponentParser => ({
  tag,
  read: (attributes) => mdxProps(attributes, { literal: ['texExpression'] }) as JsonObject,
});

/** How each Sciobot payload type reads its MDX element, keyed by payload type. */
export const SCIOBOT_PARSERS: Readonly<Record<string, MdxComponentParser>> = {
  lesson_phase: phase('phase'),
  recap_phase: phase('recap_phase'),
  lesson_activity: plain('activity'),
  column_group: plain('column_group'),
  column: plain('column'),
  lesson_title: {
    tag: 'lesson_title',
    read: (attributes) =>
      mdxProps(attributes, {
        rename: {
          organizational_minutes: 'organizationalMinutes',
          lesson_minutes: 'lessonMinutes',
        },
      }) as JsonObject,
  },
  preparation_recap: plain('preparation_recap'),
  recap_goals: plain('recap_goals'),
  recap_curriculum: plain('recap_curriculum'),
  recap_materials: plain('recap_materials'),
  recap_adaptations: plain('recap_adaptations'),
  lesson_visual: plain('lesson_visual'),
  listening: plain('listening'),
  callout,
  story: plain('story'),
  solution: plain('solution'),
  toggle: plain('toggle'),
  date: plain('date'),
  inline_equation: equation('scio-inline-equation'),
  equation: equation('scio-equation'),
  source: source('source'),
  source_excerpt: source('source'),
};
