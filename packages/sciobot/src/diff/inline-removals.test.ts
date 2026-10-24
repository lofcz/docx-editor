import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  type ImageDecodePort,
  type OoxmlPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { buildBlockIndex } from '../blocks/block-index';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
const context = { documentId: 'inline-removals', epoch: 'native' };
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const gif = Uint8Array.from(
  atob('R0lGODlhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw=='),
  (c) => c.charCodeAt(0)
);
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const kinds = {
  image: '![Cat](asset:cat "Photo")',
  math: '$x$',
  mention: '<scio-mention key="user" value="Ada" />',
  source: '<source title="Paper" url="https://source.test" />',
};
type Kind = keyof typeof kinds;
function targets(kind: Kind, ownership: string) {
  const wrap = (text: string) =>
    ownership === 'plain'
      ? text
      : `[${ownership === 'code-linked' ? `<code>${text}</code>` : text}](https://outer.test "Read")`;
  return {
    seed: wrap(`Lead. anchor ${kinds[kind]} bridge $z$ Tail.`),
    target: wrap('Lead. anchor  bridge $z$ Tail.'),
  };
}
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  const result = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target
  );
  if (!result.success) throw Error(result.error);
  const decoded = decodeEditPlan(serializeEditPlan(result.plan));
  expect(decoded).toEqual({ success: true, plan: result.plan });
  return result.plan;
}
async function apply(editor: DocxEditorInstance, target: string) {
  const planned = await plan(editor, target);
  const prepared = await prepareInlineImageResources(
    planned,
    async (reference) => ({
      bytes: reference === 'asset:dog' ? gif : png,
      mime: reference === 'asset:dog' ? 'image/gif' : 'image/png',
      widthPoints: 12,
      heightPoints: 8,
    }),
    decode
  );
  if (!prepared.success) throw Error(prepared.error);
  const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
  if (!result.success) throw Error(result.error);
  return planned;
}
function native(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes]
    .filter(([name]) => /(^|\/)media\//.test(name))
    .sort(([a], [b]) => a.localeCompare(b));
function snapshot(editor: DocxEditorInstance) {
  const pkg = editor.surface!.session.currentPackage(),
    index = buildBlockIndex(pkg);
  expect(index.diagnostics).toEqual([]);
  return {
    pkg,
    part: editor.surface!.session.part(),
    word: canonicalOoxmlFingerprint(editor.surface!.session.part()),
    media: media(pkg),
    payloads: [...index.byId]
      .map(([id, b]) => [id, b.payload])
      .sort(([a], [b]) => String(a).localeCompare(String(b))),
  };
}
function expectedRemoval(before: ReturnType<typeof snapshot>, controlId: string) {
  return native(
    native(before.part, { op: 'setContentControlProperties', controlId, lock: 'unlocked' }),
    { op: 'removeContentControl', controlId, keepContent: false }
  );
}
function verify(before: ReturnType<typeof snapshot>, editor: DocxEditorInstance, kind: Kind) {
  const current = snapshot(editor),
    old = [...buildBlockIndex(before.pkg).byId.values()];
  const removed = old.find(
    (b) =>
      b.payload.type ===
      (kind === 'math'
        ? 'inline_equation'
        : kind === 'image'
          ? 'inline_image'
          : kind === 'mention'
            ? 'mention'
            : 'source')
  );
  if (!removed) throw Error(`Missing ${kind} source`);
  expect(current.word).toBe(canonicalOoxmlFingerprint(expectedRemoval(before, removed.controlId)));
  expect(findNode(current.part, removed.controlId)).toBeNull();
  expect(current.payloads).toEqual(before.payloads.filter(([id]) => id !== removed.id));
  expect(current.media).toEqual(kind === 'image' ? [] : before.media);
  for (const block of old.filter((b) => b.id !== removed.id))
    expect(canonicalOoxmlFingerprint(findNode(current.part, block.controlId)!)).toBe(
      canonicalOoxmlFingerprint(findNode(before.part, block.controlId)!)
    );
  return current;
}
for (const kind of Object.keys(kinds) as Kind[])
  for (const ownership of ['plain', 'linked', 'code-linked'])
    test(`native inline removal preserves surviving identities: ${kind}/${ownership}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        imageDecodePort: decode,
        document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
      });
      const target = targets(kind, ownership);
      try {
        await apply(editor, target.seed);
        const before = snapshot(editor),
          planned = await apply(editor, target.target);
        expect(planned.operations.map((op) => op.kind)).toEqual(['inline-remove']);
        const after = verify(before, editor, kind);
        expect((await plan(editor, target.target)).operations).toEqual([]);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        const undone = snapshot(editor);
        expect([undone.word, undone.payloads, undone.media]).toEqual([
          before.word,
          before.payloads,
          before.media,
        ]);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        const redone = snapshot(editor);
        expect([redone.word, redone.payloads, redone.media]).toEqual([
          after.word,
          after.payloads,
          after.media,
        ]);
        const saved = new Uint8Array(await editor.save()),
          loaded = readOoxmlPackage(saved);
        if (!loaded.ok) throw Error(loaded.reason);
        expect(
          canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
        ).toBe(after.word);
        const reopened = createDocxEditor({
          container: document.createElement('div'),
          imageDecodePort: decode,
          document: saved,
        });
        try {
          expect((await plan(reopened, target.target)).operations).toEqual([]);
        } finally {
          reopened.destroy();
        }
      } finally {
        editor.destroy();
      }
    });
for (const kind of Object.keys(kinds) as Kind[])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    test(`native removal preserves peer typing and both histories: ${kind}/${clients}`, async () => {
      const target = targets(kind, 'code-linked');
      const seed = createDocxEditor({
        container: document.createElement('div'),
        imageDecodePort: decode,
        document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
      });
      let documentBytes: Uint8Array;
      try {
        await apply(seed, target.seed);
        documentBytes = new Uint8Array(await seed.save());
      } finally {
        seed.destroy();
      }
      const options = {
        document: documentBytes,
        ownedAttributes: true,
        clientIds: clients,
        editor: { imageDecodePort: decode },
      };
      const pair = await createTwoEditorFixture(options),
        actor = pair.editors[0]!,
        peer = pair.editors[1]!;
      try {
        const before = snapshot(actor),
          paragraphId = actor.surface!.session.paragraphIds()[0]!;
        pair.pause();
        await apply(actor, target.target);
        const after = verify(before, actor, kind);
        expect(
          peer.surface!.applyAutomationOps(
            () => [{ op: 'insertText', paragraphId, offset: 1, text: ' peer' }],
            { kind: 'body' },
            [],
            false,
            true
          ).rejected
        ).toBe(false);
        pair.reconnect(0);
        await pair.settle();
        const check = (active: boolean, typed: boolean) => {
          const state = active ? after : before;
          const word = typed
            ? canonicalOoxmlFingerprint(
                native(state.part, { op: 'insertText', paragraphId, offset: 1, text: ' peer' })
              )
            : state.word;
          for (const editor of pair.editors) {
            const current = snapshot(editor);
            expect([current.word, current.payloads, current.media]).toEqual([
              word,
              state.payloads,
              state.media,
            ]);
          }
        };
        check(true, true);
        expect(actor.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(false, true);
        expect(actor.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(true, true);
        expect(peer.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(true, false);
        expect(actor.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(false, false);
        expect(peer.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(false, true);
        expect(actor.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(true, true);
        peer.surface!.setSelection({
          anchor: { paragraphId, offset: 1 },
          head: { paragraphId, offset: 1 },
        });
        await pair.settle();
        const geometry = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId, offset: 1 })
        );
        expect(geometry[0]).not.toBeNull();
        expect(geometry[1]).not.toBeNull();
        expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
          x: geometry[1]!.x,
          y: geometry[1]!.y,
        });
        for (const shared of [false, true]) {
          const reopened = await createTwoEditorFixture({
            ...options,
            clientIds: undefined,
            document: new Uint8Array(await actor.save()),
            ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
          });
          try {
            for (const editor of reopened.editors) {
              const state = snapshot(editor);
              expect([state.word, state.payloads, state.media]).toEqual([
                snapshot(actor).word,
                after.payloads,
                after.media,
              ]);
            }
          } finally {
            reopened.destroy();
          }
        }
      } finally {
        pair.destroy();
      }
    });

for (const owner of ['paragraph', 'cell', 'container'])
  test(`removal composes text, births, and image replacement in a retained ${owner}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      imageDecodePort: decode,
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    const wrap = (text: string) =>
      owner === 'cell'
        ? `| Value |\n| --- |\n| ${text} |`
        : owner === 'container'
          ? `<activity name="Read" duration="1">\n\n${text}\n\n</activity>`
          : text;
    const seed = wrap(
      'Lead. ![Cat](asset:cat "Photo") anchor $x$ bridge <scio-mention key="user" value="Ada" /> <source title="Paper" url="https://source.test" /> end ![Cat](asset:cat "Photo") Tail.'
    );
    const target = wrap(
      'Lead. fresh $y$ anchor  bridge <scio-mention key="user" value="Ada" /> <source title="Changed" url="https://source.test" /> end ![Dog](asset:dog "Changed") Tail.'
    );
    try {
      await apply(editor, seed);
      const before = snapshot(editor),
        old = [...buildBlockIndex(before.pkg).byId.values()];
      const result = await apply(editor, target);
      expect(result.operations.filter((op) => op.kind === 'inline-remove')).toHaveLength(2);
      expect(result.operations.filter((op) => op.kind === 'inline-image')).toHaveLength(1);
      expect(result.operations.filter((op) => op.kind === 'inline-atoms')).toHaveLength(1);
      const after = snapshot(editor),
        next = buildBlockIndex(after.pkg);
      const mention = old.find((b) => b.payload.type === 'mention')!;
      expect(next.byId.get(mention.id)?.payload).toEqual(mention.payload);
      expect(canonicalOoxmlFingerprint(findNode(after.part, mention.controlId)!)).toBe(
        canonicalOoxmlFingerprint(findNode(before.part, mention.controlId)!)
      );
      const source = old.find((b) => b.payload.type === 'source')!;
      expect(next.byId.get(source.id)?.controlId).toBe(source.controlId);
      expect(next.byId.get(source.id)?.payload.attributes.title).toBe('Changed');
      expect(canonicalOoxmlFingerprint(findNode(after.part, source.controlId)!)).toBe(
        canonicalOoxmlFingerprint(findNode(before.part, source.controlId)!)
      );
      const retainedImage = old.filter((b) => b.payload.type === 'inline_image')[1]!;
      expect(next.byId.get(retainedImage.id)?.controlId).toBe(retainedImage.controlId);
      expect(next.byId.get(retainedImage.id)?.payload.attributes.reference).toBe('asset:dog');
      expect((await plan(editor, target)).operations).toEqual([]);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      const undone = snapshot(editor);
      expect([undone.word, undone.payloads, undone.media]).toEqual([
        before.word,
        before.payloads,
        before.media,
      ]);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const redone = snapshot(editor);
      expect([redone.word, redone.payloads, redone.media]).toEqual([
        after.word,
        after.payloads,
        after.media,
      ]);
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        imageDecodePort: decode,
        document: new Uint8Array(await editor.save()),
      });
      try {
        expect((await plan(reopened, target)).operations).toEqual([]);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

test('native removal refuses forged payloads, identities, and subtrees before mutation', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  const target = targets('image', 'linked');
  try {
    await apply(editor, target.seed);
    const before = snapshot(editor),
      planned = await plan(editor, target.target);
    const removal = planned.operations[0]!;
    expect(removal.kind).toBe('inline-remove');
    if (removal.kind !== 'inline-remove') throw Error('Missing removal');
    for (const forged of [
      { ...removal, fingerprint: 'forged' },
      { ...removal, controlId: 'missing' },
      { ...removal, paragraphId: 'missing' },
      {
        ...removal,
        before: { ...removal.before, attributes: { ...removal.before.attributes, alt: 'Wrong' } },
      },
    ]) {
      const result = applyDirectPlan(editor, { ...planned, operations: [forged] }, context);
      expect(result.success).toBe(false);
      const current = snapshot(editor);
      expect([current.word, current.payloads, current.media]).toEqual([
        before.word,
        before.payloads,
        before.media,
      ]);
    }
    for (const forged of [
      { ...removal, before: { ...removal.before, type: 'activity' } },
      { ...removal, blockId: 'another-id' },
      { ...removal, fingerprint: '' },
      { ...removal, unknown: true },
    ]) {
      expect(decodeEditPlan(JSON.stringify({ ...planned, operations: [forged] })).success).toBe(
        false
      );
    }
  } finally {
    editor.destroy();
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`combined removal and birth preserve both native histories: ${clients}`, async () => {
    const seedText =
      '[<code>Lead. ![Cat](asset:cat "Photo") anchor $x$ bridge $z$ Tail.</code>](https://outer.test "Read")';
    const target =
      '[<code>Lead. fresh <scio-mention key="new-user" value="Ada" /> anchor  bridge $z$ Tail.</code>](https://outer.test "Read")';
    const seed = createDocxEditor({
      container: document.createElement('div'),
      imageDecodePort: decode,
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    let documentBytes: Uint8Array;
    try {
      await apply(seed, seedText);
      documentBytes = new Uint8Array(await seed.save());
    } finally {
      seed.destroy();
    }
    const options = {
      document: documentBytes,
      ownedAttributes: true,
      clientIds: clients,
      editor: { imageDecodePort: decode },
    };
    const pair = await createTwoEditorFixture(options),
      actor = pair.editors[0]!,
      peer = pair.editors[1]!;
    try {
      const before = snapshot(actor),
        paragraphId = actor.surface!.session.paragraphIds()[0]!;
      pair.pause();
      const planned = await apply(actor, target);
      expect(planned.operations.filter((op) => op.kind === 'inline-remove')).toHaveLength(2);
      expect(planned.operations.filter((op) => op.kind === 'inline-atoms')).toHaveLength(1);
      const after = snapshot(actor),
        old = buildBlockIndex(before.pkg),
        next = buildBlockIndex(after.pkg);
      const retained = [...old.byId.values()].find(
        (block) =>
          block.payload.type === 'inline_equation' && block.payload.attributes.texExpression === 'z'
      )!;
      expect(next.byId.get(retained.id)?.payload).toEqual(retained.payload);
      expect(canonicalOoxmlFingerprint(findNode(after.part, retained.controlId)!)).toBe(
        canonicalOoxmlFingerprint(findNode(before.part, retained.controlId)!)
      );
      expect(after.media).toEqual([]);
      expect((await plan(actor, target)).operations).toEqual([]);
      expect(
        peer.surface!.applyAutomationOps(
          () => [{ op: 'insertText', paragraphId, offset: 1, text: ' peer' }],
          { kind: 'body' },
          [],
          false,
          true
        ).rejected
      ).toBe(false);
      pair.reconnect(0);
      await pair.settle();
      const check = (active: boolean, typed: boolean) => {
        const state = active ? after : before;
        const word = typed
          ? canonicalOoxmlFingerprint(
              native(state.part, { op: 'insertText', paragraphId, offset: 1, text: ' peer' })
            )
          : state.word;
        for (const editor of pair.editors) {
          const current = snapshot(editor);
          expect([current.word, current.payloads, current.media]).toEqual([
            word,
            state.payloads,
            state.media,
          ]);
        }
      };
      check(true, true);
      expect(actor.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      check(false, true);
      expect(actor.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      check(true, true);
      expect(peer.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      check(true, false);
      expect(actor.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      check(false, false);
      expect(peer.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      check(false, true);
      expect(actor.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      check(true, true);
      for (const shared of [false, true]) {
        const reopened = await createTwoEditorFixture({
          ...options,
          clientIds: undefined,
          document: new Uint8Array(await actor.save()),
          ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
        });
        try {
          for (const editor of reopened.editors) {
            const current = snapshot(editor);
            expect([current.word, current.payloads, current.media]).toEqual([
              snapshot(actor).word,
              after.payloads,
              after.media,
            ]);
          }
        } finally {
          reopened.destroy();
        }
      }
    } finally {
      pair.destroy();
    }
  });

test('shared removal requires complete ownership before publication', async () => {
  const seed = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decode,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  const target = targets('image', 'plain');
  let documentBytes: Uint8Array;
  try {
    await apply(seed, target.seed);
    documentBytes = new Uint8Array(await seed.save());
  } finally {
    seed.destroy();
  }
  const pair = await createTwoEditorFixture({
    document: documentBytes,
    editor: { imageDecodePort: decode },
  });
  try {
    const before = pair.editors.map(snapshot),
      planned = await plan(pair.editors[0]!, target.target);
    const vectors = pair.documents.map((doc) => Y.encodeStateVector(doc));
    const result = applyDirectPlan(pair.editors[0]!, planned, context);
    expect(result.success).toBe(false);
    if (!result.success) expect(result.code).toBe('unsupported-operation');
    await pair.settle();
    expect(pair.documents.map((doc) => Y.encodeStateVector(doc))).toEqual(vectors);
    pair.editors.forEach((editor, index) => {
      const current = snapshot(editor),
        expected = before[index]!;
      expect([current.word, current.payloads, current.media]).toEqual([
        expected.word,
        expected.payloads,
        expected.media,
      ]);
    });
  } finally {
    pair.destroy();
  }
});
