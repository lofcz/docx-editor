import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  hostControl,
  hostParagraph,
  reactBlockDocument,
} from '../../test/support/react-block-document';
import { docxFixture } from '../../test/support/document';
import { writeBlockPayload } from './payload-part';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import type { JsonObject } from './payload-schema';
import { writeOoxmlPackage } from '@docx-editor.dev/core/store';

const CLIENTS = [
  [73, 41],
  [41, 73],
] as const;
const RECEIVERS = [0, 1] as const;

function controlOf(editor: DocxEditorInstance, id: string): OoxmlNode {
  const visit = (node: OoxmlNode): OoxmlNode | null => {
    if (node.kind === 'textValue') return null;
    if (node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === `scio:1:${id}`)
      return node;
    for (const child of node.children) {
      const found = visit(child);
      if (found) return found;
    }
    return null;
  };
  const found = visit(editor.surface!.session.part().root);
  if (!found) throw new Error(`Missing control ${id}`);
  return found;
}

function attributesOf(editor: DocxEditorInstance, id: string): JsonObject | string {
  const store = readPayloadStore(editor.surface!.session.currentPackage());
  if (store.status !== 'ready') return store.status === 'invalid' ? store.reason : store.status;
  const entry = store.entries.get(id)?.decoded;
  return entry?.status === 'supported' ? entry.value.attributes : String(entry?.status);
}

function write(
  editor: DocxEditorInstance,
  id: string,
  attributes: JsonObject,
  options: { label?: string; fallback?: string } = {}
) {
  const control = controlOf(editor, id);
  const first = (control.kind !== 'textValue' && findFirstParagraph(control)) || null;
  return writeBlockPayload(editor, {
    controlId: control.id,
    payload: { id, type: 'test_counter', version: 1, attributes },
    label: options.label ?? 'Word fallback',
    fallbackOps:
      options.fallback === undefined || !first
        ? []
        : [
            {
              op: 'deleteText',
              paragraphId: first,
              start: 0,
              end: paragraphTextOf(editor.surface!.session.part(), first)!.length,
            },
            { op: 'insertText', paragraphId: first, offset: 0, text: options.fallback },
          ],
  });
}

function findFirstParagraph(node: OoxmlNode): string | null {
  if (node.kind === 'textValue') return null;
  if (node.kind === 'paragraph') return node.id;
  for (const child of node.children) {
    const found = findFirstParagraph(child);
    if (found) return found;
  }
  return null;
}

async function expectConverged(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  id: string,
  expected?: JsonObject
) {
  const [alice, bob] = pair.editors;
  for (const room of pair.rooms) expect(room.session.status()).not.toBe('error');
  const left = attributesOf(alice!, id);
  expect(attributesOf(bob!, id)).toEqual(left);
  if (expected) expect(left).toEqual(expected);
  const [savedAlice, savedBob] = await pair.savedDigests();
  expect(savedAlice).toEqual(savedBob);
  return left;
}

async function reopenedAttributes(editor: DocxEditorInstance, id: string) {
  const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!opened.ok) throw new Error(opened.reason);
  const store = readPayloadStore(opened.package);
  if (store.status !== 'ready') throw new Error(store.status);
  const entry = store.entries.get(id)?.decoded;
  return entry?.status === 'supported' ? entry.value.attributes : null;
}

for (const clientIds of CLIENTS)
  for (const first of RECEIVERS)
    test(`disjoint attribute writes to one block merge; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({ document: reactBlockDocument(), clientIds });
      try {
        const [alice, bob] = pair.editors;
        pair.pause();
        expect(write(alice!, 'counter-1', { name: 'Initial', grade: 'A' }).success).toBe(true);
        expect(write(bob!, 'counter-1', { name: 'Initial', lessonMinutes: 45 }).success).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        const merged = { name: 'Initial', grade: 'A', lessonMinutes: 45 };
        await expectConverged(pair, 'counter-1', merged);
        expect(await reopenedAttributes(alice!, 'counter-1')).toEqual(merged);

        // Either write undoes alone and redoes, and both sessions follow.
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await expectConverged(pair, 'counter-1', { name: 'Initial', lessonMinutes: 45 });
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await expectConverged(pair, 'counter-1', { name: 'Initial' });
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await expectConverged(pair, 'counter-1', { name: 'Initial', grade: 'A' });
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await expectConverged(pair, 'counter-1', merged);

        // A reopened room continues from the converged record.
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await bob!.save()),
        });
        try {
          for (const editor of reopened.editors)
            expect(attributesOf(editor, 'counter-1')).toEqual(merged);
          expect(write(reopened.editors[0]!, 'counter-1', { ...merged, grade: 'B' }).success).toBe(
            true
          );
          await reopened.settle();
          for (const editor of reopened.editors)
            expect(attributesOf(editor, 'counter-1')).toEqual({ ...merged, grade: 'B' });
        } finally {
          reopened.destroy();
        }
      } finally {
        pair.destroy();
      }
    }, 60_000);

for (const clientIds of CLIENTS)
  for (const first of RECEIVERS)
    test(`conflicting writes to one attribute pick one winner everywhere; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({ document: reactBlockDocument(), clientIds });
      try {
        const [alice, bob] = pair.editors;
        pair.pause();
        // Same key, a removal racing a change, and one disjoint key on each side.
        expect(
          write(alice!, 'counter-1', { name: 'Warmup', grade: 'A' }, { label: 'Warmup' }).success
        ).toBe(true);
        expect(
          write(bob!, 'counter-1', { name: 'Cooldown', minutes: 10 }, { label: 'Cooldown' }).success
        ).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        const merged = (await expectConverged(pair, 'counter-1')) as JsonObject;
        // A register winner, never a character merge of both values.
        expect(['Warmup', 'Cooldown']).toContain(merged.name as string);
        expect(merged).toMatchObject({ grade: 'A', minutes: 10 });

        // The next write to the contested key leaves one register for it.
        expect(write(bob!, 'counter-1', { ...merged, name: 'Final' }).success).toBe(true);
        await pair.settle();
        await expectConverged(pair, 'counter-1', { ...merged, name: 'Final' });
        const store = readPayloadStore(bob!.surface!.session.currentPackage());
        if (store.status !== 'ready') throw new Error(store.status);
        const record = (store.part &&
          bob!
            .surface!.session.currentPackage()
            .parts.get(store.part.partName)!
            .root.children.find(
              (child) =>
                child.kind !== 'textValue' &&
                child.attributes.some((a) => a.localName === 'id' && a.value === 'scio_counter-1')
            )) as Exclude<OoxmlNode, { kind: 'textValue' }>;
        const nameRegisters = record.children.filter(
          (child) =>
            child.kind !== 'textValue' &&
            child.localName === 'attr' &&
            child.attributes.some((a) => a.localName === 'name' && a.value === 'name')
        );
        expect(nameRegisters).toHaveLength(1);
      } finally {
        pair.destroy();
      }
    }, 60_000);

for (const clientIds of CLIENTS)
  test(`removing an attribute while a peer changes it converges; clients ${clientIds.join('/')}`, async () => {
    const pair = await createTwoEditorFixture({ document: reactBlockDocument(), clientIds });
    try {
      const [alice, bob] = pair.editors;
      expect(write(alice!, 'counter-1', { name: 'Initial', grade: 'A' }).success).toBe(true);
      await pair.settle();
      pair.pause();
      expect(write(alice!, 'counter-1', { name: 'Initial' }).success).toBe(true);
      expect(write(bob!, 'counter-1', { name: 'Initial', grade: 'B', extra: true }).success).toBe(
        true
      );
      pair.reconnect(0);
      await pair.settle();
      const merged = (await expectConverged(pair, 'counter-1')) as JsonObject;
      expect(merged.extra).toBe(true);
      expect([undefined, 'B']).toContain(merged.grade as string | undefined);
    } finally {
      pair.destroy();
    }
  }, 60_000);

test('writes to different blocks still merge, with readable fallback text', async () => {
  const pair = await createTwoEditorFixture({
    document: reactBlockDocument(),
    clientIds: [73, 41],
  });
  try {
    const [alice, bob] = pair.editors;
    pair.pause();
    expect(
      write(alice!, 'counter-1', { name: 'Initial', grade: 'A' }, { fallback: 'Grade A' }).success
    ).toBe(true);
    expect(write(bob!, 'future-1', { name: 'Initial', minutes: 'B' }).success).toBe(true);
    pair.reconnect(1);
    await pair.settle();
    await expectConverged(pair, 'counter-1', { name: 'Initial', grade: 'A' });
    await expectConverged(pair, 'future-1', { name: 'Initial', minutes: 'B' });
    for (const editor of pair.editors)
      expect(
        paragraphTextOf(
          editor.surface!.session.part(),
          findFirstParagraph(controlOf(editor, 'counter-1'))!
        )
      ).toBe('Grade A');
  } finally {
    pair.destroy();
  }
}, 60_000);

test('a payload write during peer deletion of the block leaves both sessions equal', async () => {
  const pair = await createTwoEditorFixture({
    document: reactBlockDocument(),
    clientIds: [41, 73],
  });
  try {
    const [alice, bob] = pair.editors;
    pair.pause();
    const control = controlOf(alice!, 'counter-1');
    expect(
      alice!.surface!.applyAutomationOps(() => [
        { op: 'removeContentControl', controlId: control.id, keepContent: false },
      ]).rejected
    ).toBe(false);
    expect(write(bob!, 'counter-1', { name: 'Initial', grade: 'B' }).success).toBe(true);
    pair.reconnect(0);
    await pair.settle();
    for (const room of pair.rooms) expect(room.session.status()).not.toBe('error');
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
    expect(attributesOf(alice!, 'counter-1')).toEqual(attributesOf(bob!, 'counter-1'));
  } finally {
    pair.destroy();
  }
}, 60_000);

test('concurrent writes keep an existing alias as one property and never add a missing one', async () => {
  const read = readOoxmlPackage(
    docxFixture(
      hostParagraph('Intro') +
        hostControl('counter-1', hostParagraph('Widget fallback')).replace(
          '<w:sdtPr>',
          '<w:sdtPr><w:alias w:val="Old"/>'
        ) +
        hostControl('future-1', hostParagraph('No alias'))
    )
  );
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const id of ['counter-1', 'future-1']) {
    const stored = withPayloadRecord(
      pkg,
      { id, type: 'test_counter', version: 1, attributes: { name: 'Initial' } },
      'Old'
    );
    if (!stored.ok) throw new Error(stored.reason);
    pkg = stored.pkg;
  }
  const pair = await createTwoEditorFixture({
    document: writeOoxmlPackage(pkg),
    clientIds: [73, 41],
  });
  try {
    const [alice, bob] = pair.editors;
    pair.pause();
    for (const id of ['counter-1', 'future-1']) {
      expect(write(alice!, id, { name: 'A' }, { label: 'Alice label' }).success).toBe(true);
      expect(write(bob!, id, { name: 'B' }, { label: 'Bob label' }).success).toBe(true);
    }
    pair.reconnect(0);
    await pair.settle();
    for (const room of pair.rooms) expect(room.session.status()).not.toBe('error');
    const aliases = pair.editors.map(
      (editor) => contentControlPropertiesOf(controlOf(editor, 'counter-1')).alias
    );
    expect(aliases[0]).toBe(aliases[1]);
    expect(['Alice label', 'Bob label']).toContain(aliases[0]!);
    for (const editor of pair.editors) {
      expect(contentControlPropertiesOf(controlOf(editor, 'future-1')).alias).toBeUndefined();
      const sdtPr = (controlOf(editor, 'counter-1') as Exclude<OoxmlNode, { kind: 'textValue' }>)
        .children[0] as Exclude<OoxmlNode, { kind: 'textValue' }>;
      expect(
        sdtPr.children.filter((child) => child.kind !== 'textValue' && child.localName === 'alias')
      ).toHaveLength(1);
      expect(
        findNode(editor.surface!.session.part(), controlOf(editor, 'future-1').id)
      ).toBeTruthy();
    }
    await expectConverged(pair, 'counter-1');
  } finally {
    pair.destroy();
  }
}, 60_000);
