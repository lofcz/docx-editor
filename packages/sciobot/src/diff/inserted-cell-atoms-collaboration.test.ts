import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { table } from '../../test/support/target-table-fixture';
import { prepareNativeDocument } from '../native/prepare-document';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { buildBlockIndex } from '../blocks/block-index';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { parseTargetMarkdown } from '../projection/parse-target';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
import { targetAstKey } from './target-key';
const context = { documentId: 'shared-cell-atoms', epoch: 'native' };
const atoms =
  '$x$ ![Cat](asset:cat "Picture") [Ada](mention:user) <source url="https://example.test" title="Paper" />';
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decode = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => /(^|\/)media\//.test(name));
const key = (text: string) => {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
};
for (const kind of ['row', 'column'] as const)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`new ${kind} atoms retain peer cell text through both histories, geometry, Word and room reopen: ${clientIds}/${receiver}`, async () => {
        const prepared = prepareNativeDocument(
          docxFixture(
            table([
              ['A', 'B'],
              ['Keep', 'Peer'],
            ])
          )
        );
        if (!prepared.success) throw Error(prepared.error);
        const fixture = {
          document: prepared.document,
          clientIds,
          editor: { imageDecodePort: decode },
        };
        const pair = await createTwoEditorFixture(fixture);
        try {
          const [alice, bob] = pair.editors;
          const peer = bob!
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Peer')!;
          expect(peer).toBeDefined();
          const target =
            kind === 'row'
              ? `| A | B |\n| - | - |\n| Keep | Peer |\n| ${atoms} | Tail |`
              : `| A | New | B |\n| - | - | - |\n| Keep | ${atoms} | Peer |`;
          const read = async (editor: typeof alice) => {
            const revision = readEditorRevision(editor!, context.documentId, context.epoch)!;
            const source = await readWorkingView(editor!.surface!.session, revision);
            if (!source.success) throw Error(source.error);
            return {
              revision,
              pkg: editor!.surface!.session.currentPackage(),
              projection: source.view.projection,
            };
          };
          const check = async (born = true, peerEdit = true) => {
            const updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
            for (const editor of pair.editors) {
              const pkg = editor.surface!.session.currentPackage();
              const projection = await read(editor);
              const expected = born ? target : '| A | B |\n| - | - |\n| Keep | Peer |';
              expect(key(projection.projection.text)).toBe(
                key(peerEdit ? expected.replace('Peer', 'Peer edit') : expected)
              );
              expect(paragraphTextOf(editor.surface!.session.part(), peer)).toBe(
                peerEdit ? 'Peer edit' : 'Peer'
              );
              const index = buildBlockIndex(pkg);
              expect(index.diagnostics).toEqual([]);
              expect([...index.byId.values()].map((entry) => entry.payload.type).sort()).toEqual(
                born ? ['inline_equation', 'inline_image', 'mention', 'source'] : []
              );
              expect(media(pkg)).toHaveLength(born ? 1 : 0);
              expect(
                (pkg.relationships.get(pkg.mainDocumentPart) ?? []).filter((entry) =>
                  entry.type.endsWith('/image')
                )
              ).toHaveLength(born ? 1 : 0);
              if (born) expect(media(pkg)[0]![1]).toEqual(png);
              expect(editor.surface!.session.currentPackage()).toBe(pkg);
            }
            pair.documents.forEach((doc, index) =>
              expect(Y.encodeStateAsUpdate(doc)).toEqual(updates[index]!)
            );
            expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(bob!.surface!.session.part())
            );
            expect(media(alice!.surface!.session.currentPackage())).toEqual(
              media(bob!.surface!.session.currentPackage())
            );
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
          };
          const planned = planTargetDocument(await read(alice), target);
          if (!planned.success) throw Error(planned.error);
          const resources = await prepareInlineImageResources(
            planned.plan,
            async (ref) => {
              expect(ref).toBe('asset:cat');
              return { bytes: png, mime: 'image/png' };
            },
            decode
          );
          if (!resources.success) throw Error(resources.error);
          pair.pause();
          const applied = applyDirectPlan(alice!, planned.plan, {
            ...context,
            images: resources.images,
          });
          if (!applied.success) throw Error(applied.error);
          expect(
            bob!.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId: peer, offset: 4, text: ' edit' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          await check();
          const expectedWord = canonicalOoxmlFingerprint(alice!.surface!.session.part());
          const expectedPayloads = [
            ...buildBlockIndex(alice!.surface!.session.currentPackage()).byId.values(),
          ].map((entry) => entry.payload);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(false, true);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(false, false);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(false, true);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(expectedWord);
          bob!.surface!.setSelection({
            anchor: { paragraphId: peer, offset: 9 },
            head: { paragraphId: peer, offset: 9 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(peer);
          expect(remote?.head.offset).toBe(9);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: peer, offset: 9 })
          );
          expect(carets.every(Boolean)).toBe(true);
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          for (const shared of [false, true]) {
            const reopened = await createTwoEditorFixture({
              ...fixture,
              clientIds: undefined,
              document: new Uint8Array(await alice!.save()),
              ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              await reopened.settle();
              for (const editor of reopened.editors) {
                expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                  expectedWord
                );
                expect(media(editor.surface!.session.currentPackage())).toEqual(
                  media(alice!.surface!.session.currentPackage())
                );
                expect(
                  [...buildBlockIndex(editor.surface!.session.currentPackage()).byId.values()].map(
                    (entry) => entry.payload
                  )
                ).toEqual(expectedPayloads);
              }
            } finally {
              reopened.destroy();
            }
          }
        } finally {
          pair.destroy();
        }
      }, 60_000);

test('shared new-cell atoms require prepared payload storage before any publication or history change', async () => {
  const pair = await createTwoEditorFixture({
    document: docxFixture(
      table([
        ['A', 'B'],
        ['Keep', 'Peer'],
      ])
    ),
  });
  try {
    const editor = pair.editors[0]!;
    const pkg = editor.surface!.session.currentPackage();
    const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
    const source = await readWorkingView(editor.surface!.session, revision);
    if (!source.success) throw Error(source.error);
    const planned = planTargetDocument(
      { revision, pkg, projection: source.view.projection },
      '| A | B |\n| - | - |\n| Keep | Peer |\n| $x$ | Tail |'
    );
    if (!planned.success) throw Error(planned.error);
    const updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    const word = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const result = applyDirectPlan(editor, planned.plan, context);
    expect(result).toMatchObject({
      success: false,
      code: 'unsupported-operation',
      error: 'Prepare native payload storage before sharing inline atom births.',
    });
    expect(editor.surface!.session.currentPackage()).toBe(pkg);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(word);
    expect(editor.surface!.session.canUndo()).toBe(false);
    pair.documents.forEach((doc, index) =>
      expect(Y.encodeStateAsUpdate(doc)).toEqual(updates[index]!)
    );
  } finally {
    pair.destroy();
  }
});
