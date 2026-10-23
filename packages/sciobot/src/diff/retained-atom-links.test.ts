import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  ensureHyperlinkRelationship,
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
const context = { documentId: 'retained-links', epoch: 'native' };
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const main = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
function descendants(node: OoxmlNode): OoxmlNode[] {
  return [node, ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants))];
}
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  let n = 0;
  const result = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target,
    { allocateBlockId: () => `retained-link-${++n}` }
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
type Link = { start: number; end: number; url: string; title: string };
/** Independent native envelope oracle, with explicit expected spans and no target parser. */
function nativeLinks(
  pkg: OoxmlPackage,
  paragraphId: string,
  wanted: readonly Link[]
): OoxmlPackage {
  let part = main(pkg);
  const paragraph = findNode(part, paragraphId)!;
  if (paragraph.kind !== 'paragraph') throw Error('Missing paragraph');
  const offsets = paragraphOffsetIndex(paragraph);
  const links = descendants(paragraph).filter((node) => node.kind === 'hyperlink');
  const kept = new Set<string>();
  const apply = (op: TreeDocOp) => {
    const r = applyTreeOp(part, op);
    if (!r.ok) throw Error(r.reason);
    part = r.part;
  };
  for (const old of links) {
    const span = offsets.spanOf(old);
    if (!wanted.some((next) => span?.start === next.start && span.end === next.end))
      apply({ op: 'removeHyperlink', linkId: old.id });
  }
  for (const next of wanted) {
    const match = links.find(
      (old) =>
        !kept.has(old.id) &&
        offsets.spanOf(old)?.start === next.start &&
        offsets.spanOf(old)?.end === next.end
    );
    const anchor = next.url.startsWith('#') ? next.url.slice(1) : null;
    let relationshipId: string | undefined;
    if (anchor === null) {
      const made = ensureHyperlinkRelationship(withPart(pkg, part), next.url, pkg.mainDocumentPart);
      if (!made) throw Error('Invalid oracle URL');
      pkg = made.pkg;
      relationshipId = made.relationshipId;
    }
    if (match) {
      kept.add(match.id);
      apply({
        op: 'setHyperlinkTarget',
        linkId: match.id,
        ...(anchor === null ? { relationshipId } : { anchor }),
        tooltip: next.title,
      });
    } else
      apply({
        op: 'insertHyperlink',
        paragraphId,
        start: next.start,
        end: next.end,
        ...(anchor === null ? { relationshipId } : { anchor }),
        tooltip: next.title,
      });
  }
  return withPart(pkg, part);
}
for (const atom of [
  '$x$',
  '[Ada](mention:user-7)',
  '<source title="Paper" url="https://source.test" />',
  '![Cat](asset:cat "Photo")',
])
  test(`retained atom link envelopes preserve native content through every ownership change: ${atom}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t>Keep.</w:t></w:r></w:p>'
      ),
      imageDecodePort: decode,
    });
    try {
      const plain = `Keep. A ${atom} Z`,
        birth = await plan(editor, plain);
      const prepared = await prepareInlineImageResources(
        birth,
        async () => ({ bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 }),
        decode
      );
      if (!prepared.success) throw Error(prepared.error);
      const seeded = applyDirectPlan(editor, birth, { ...context, images: prepared.images });
      if (!seeded.success) throw Error(seeded.error);
      const original = editor.surface!.session.currentPackage(),
        blocks = buildBlockIndex(original);
      const block = [...blocks.byId.values()][0]!,
        paragraphId = editor.surface!.session.paragraphIds()[0]!;
      const paragraph = findNode(main(original), paragraphId)!;
      if (paragraph.kind !== 'paragraph') throw Error('Missing paragraph');
      const span = paragraphOffsetIndex(paragraph).spanOf(
        findNode(main(original), block.controlId)!
      )!;
      const atomWord = canonicalOoxmlFingerprint(findNode(main(original), block.controlId)!);
      const end = paragraphTextOf(main(original), paragraphId)!.length;
      const imageParts = [...original.partBytes].filter(([name]) => name.includes('/media/'));
      const linkedAtom = atom.startsWith('[Ada]')
        ? '<scio-mention key="user-7" value="Ada" />'
        : atom;
      const phases: [string, Link[]][] = [
        [
          `Keep. A [${linkedAtom}](https://first.test "Read") Z`,
          [{ ...span, url: 'https://first.test', title: 'Read' }],
        ],
        [
          `Keep. [A ${linkedAtom} Z](https://first.test "Read")`,
          [{ start: 6, end, url: 'https://first.test', title: 'Read' }],
        ],
        [
          `Keep. [A ${linkedAtom} Z](https://second.test "Changed")`,
          [{ start: 6, end, url: 'https://second.test', title: 'Changed' }],
        ],
        [
          `Keep. A [${linkedAtom}](#target "Changed") Z`,
          [{ ...span, url: '#target', title: 'Changed' }],
        ],
        [plain, []],
        [
          `Keep. [A ${linkedAtom}](https://first.test "Read") [Z](https://first.test "Read")`,
          [
            { start: 6, end: span.end, url: 'https://first.test', title: 'Read' },
            { start: end - 1, end, url: 'https://first.test', title: 'Read' },
          ],
        ],
        [plain, []],
      ];
      for (const [target, links] of phases) {
        const current = editor.surface!.session.currentPackage(),
          before = canonicalOoxmlFingerprint(main(current));
        const expected = nativeLinks(current, paragraphId, links);
        const planned = await plan(editor, target);
        expect(planned.operations.length).toBeGreaterThan(0);
        expect(
          planned.operations.every((op) => ['format', 'link', 'unlink'].includes(op.kind))
        ).toBe(true);
        expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({
          success: true,
          plan: planned,
        });
        const resources = await prepareInlineImageResources(
          planned,
          async () => {
            throw Error('Retained image must not resolve');
          },
          decode
        );
        expect(resources.success).toBe(true);
        const applied = applyDirectPlan(editor, planned, context);
        if (!applied.success) throw Error(applied.error);
        const after = editor.surface!.session.currentPackage(),
          word = canonicalOoxmlFingerprint(main(expected));
        expect(canonicalOoxmlFingerprint(main(after))).toBe(word);
        expect(canonicalOoxmlFingerprint(findNode(main(after), block.controlId)!)).toBe(atomWord);
        expect([...buildBlockIndex(after).byId]).toEqual([...blocks.byId]);
        expect([...after.partBytes].filter(([name]) => name.includes('/media/'))).toEqual(
          imageParts
        );
        expect((await plan(editor, target)).operations).toEqual([]);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(word);
        const bytes = new Uint8Array(await editor.save()),
          loaded = readOoxmlPackage(bytes);
        if (!loaded.ok) throw Error(loaded.reason);
        expect(canonicalOoxmlFingerprint(main(loaded.package))).toBe(word);
        const reopened = createDocxEditor({
          container: document.createElement('div'),
          document: bytes,
          imageDecodePort: decode,
        });
        try {
          expect((await plan(reopened, target)).operations).toEqual([]);
        } finally {
          reopened.destroy();
        }
      }
    } finally {
      editor.destroy();
    }
  });
