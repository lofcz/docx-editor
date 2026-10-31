import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';

/**
 * Quoted Sciobot component and storage names. The Sciobot registration owns them; a generic
 * adapter module that names one hard-codes a Sciobot rule another embedder cannot replace.
 */
export const APPLICATION_LITERAL =
  /(['"`])(?:lesson_(?:title|phase|activity|visual)|recap_(?:phase|goals|curriculum|materials|adaptations)|preparation_recap|callout|story|solution|code_block|inline_equation|equation|mention|inline_image|source|source_excerpt|listening|scio-[a-z-]+|scio:[^'"`]*|urn:sciobot:[^'"`]*)\1/g;

/**
 * Files that may name Sciobot components: the registration, and the codecs, serializers, and
 * tool contract of the Sciobot components it registers.
 */
const OWNERS = new RegExp(
  '^(?:' +
    [
      'registry/sciobot[\\w-]*\\.ts',
      'projection/(?:block-tags|mdx-tags|math|mention|mention-literal|target-mentions|inline-image|native-image|native-math)\\.ts',
      'blocks/(?:math-codec|mention-codec|inline-image-codec|payload-schema|payload-opc-codec)\\.ts',
      'tools/validation/(?:lesson-plan-schema|schema-engine)\\.ts',
    ].join('|') +
    ')$'
);
/** A line that uses one of these words for something else marks it with this comment. */
const NOT_A_COMPONENT = '// not-a-component:';

/** Literal counts per adapter source file, outside tests and the registration. */
export function applicationLiteralCounts(root: string): Record<string, number> {
  const counts: Record<string, number> = {};
  const visit = (directory: string) => {
    for (const entry of readdirSync(directory, { withFileTypes: true })) {
      const path = join(directory, entry.name);
      if (entry.isDirectory()) {
        if (entry.name !== '__tests__') visit(path);
        continue;
      }
      const name = relative(root, path);
      if (!/\.tsx?$/.test(name) || /\.test\.tsx?$/.test(name) || OWNERS.test(name)) continue;
      const found = readFileSync(path, 'utf8')
        .split('\n')
        .filter((line) => !line.includes(NOT_A_COMPONENT))
        .reduce((sum, line) => sum + (line.match(APPLICATION_LITERAL)?.length ?? 0), 0);
      if (found) counts[name] = found;
    }
  };
  visit(root);
  return Object.fromEntries(Object.entries(counts).sort(([a], [b]) => a.localeCompare(b)));
}
