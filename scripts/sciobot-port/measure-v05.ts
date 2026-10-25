import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { strToU8, zipSync } from 'fflate';
import { mountPaginatedSurface } from '../../packages/core/src/editor/paginated-surface.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function fixture(): Uint8Array {
  const pages: string[] = [];
  for (let page = 0; page < 100; page += 1) {
    const body = Array.from({ length: 12 }, (_, index) =>
      paragraph(`Page ${page + 1} paragraph ${index + 1}`)
    ).join('');
    pages.push(
      `${body}<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>`
    );
  }
  const xml = `<w:document xmlns:w="${W}"><w:body>${pages.join('')}</w:body></w:document>`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(xml),
  });
}

function percentile(values: readonly number[], fraction: number) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * fraction))] ?? 0;
}

const samples = {
  open: [] as number[],
  typing: [] as number[],
  review: [] as number[],
  movement: [] as number[],
};
const start = performance.now();
const container = document.createElement('div');
document.body.append(container);
const opened = mountPaginatedSurface(container, fixture(), { scale: 1 });
if (!opened.ok) throw new Error(opened.reason);
const surface = opened.surface;
const openedAt = performance.now();
surface.layout();
samples.open.push(openedAt - start);
const ids = surface.session.paragraphIds();
for (let index = 0; index < 20; index += 1) {
  const id = ids[12 * index + 3]!;
  const typed = performance.now();
  surface.session.applyTreeOps([{ op: 'insertText', paragraphId: id, offset: 0, text: 'x' }]);
  surface.layout();
  samples.typing.push(performance.now() - typed);
}
for (let index = 0; index < 20; index += 1) {
  const id = ids[12 * index + 4]!;
  const reviewed = performance.now();
  surface.session.applyTreeOps([{ op: 'insertText', paragraphId: id, offset: 0, text: 'review ' }]);
  surface.layout();
  samples.review.push(performance.now() - reviewed);
}
for (let index = 0; index < 20; index += 1) {
  const moved = performance.now();
  surface.revealParagraph(ids[ids.length - 1 - index * 12]!);
  samples.movement.push(performance.now() - moved);
}
const output = {
  fixture: { pages: 100, paragraphs: ids.length },
  environment: { runtime: 'Bun', platform: process.platform, arch: process.arch },
  samples,
  summary: Object.fromEntries(
    Object.entries(samples).map(([name, values]) => [
      name,
      {
        count: values.length,
        p50Ms: percentile(values, 0.5),
        p95Ms: percentile(values, 0.95),
        maxMs: Math.max(...values),
      },
    ])
  ),
  retainedParagraphs: ids.length,
};
console.log(JSON.stringify(output));
surface.destroy();
container.remove();
