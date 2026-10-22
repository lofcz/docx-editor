import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  bodyStoryRoot,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  canonicalOoxmlFingerprint,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { applyDirectPlan } from '../ai/apply-direct';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { planTargetDocument } from './plan-target';
import { parseTargetMarkdown } from '../projection/parse-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
const p = (text: string, quote = false) =>
  `<w:p><w:pPr>${quote ? '<w:pStyle w:val="Quote"/>' : ''}<keep:extra xmlns:keep="urn:keep" keep:v="unchanged"/></w:pPr><w:r><w:rPr><w:color w:val="123456"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
async function fixture(body: string) {
  const entries = unzipSync(docxFixture(body));
  entries['word/styles.xml'] = strToU8(
    `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/></w:style></w:styles>`
  );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
    )
  );
  entries['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(entries['word/_rels/document.xml.rels']!).replace(
      '</Relationships>',
      '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    )
  );
  const loaded = readOoxmlPackage(zipSync(entries));
  if (!loaded.ok) throw new Error(loaded.reason);
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(loaded.package),
  });

  const pkg = editor.surface!.session.currentPackage();
  const revision = readEditorRevision(editor, sourceRevision.documentId, sourceRevision.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw new Error(view.error);
  return { editor, pkg, revision, projection: view.view.projection };
}
const ast = (markdown: string) => {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.normalizedAst;
};
const cases = [
  {
    name: 'text and mark change',
    target: 'Before\n\n> First\n>\n> **Second changed**\n\nAfter',
    remaining: ['Before', 'First', 'Second changed', 'After'],
    retained: [0, 1, 2, 3],
    kinds: ['text', 'format'],
  },
  {
    name: 'paragraph insertion',
    target: 'Before\n\n> First\n>\n> New quote\n>\n> Second\n\nAfter',
    remaining: ['Before', 'First', 'New quote', 'Second', 'After'],
    retained: [0, 1, 2, 3],
    kinds: ['siblings'],
  },
  {
    name: 'paragraph removal',
    target: 'Before\n\n> First\n\nAfter',
    remaining: ['Before', 'First', 'After'],
    retained: [0, 1, 3],
    kinds: ['siblings'],
  },
  {
    name: 'leave quote',
    target: 'Before\n\n> First\n\nSecond\n\nAfter',
    remaining: ['Before', 'First', 'Second', 'After'],
    retained: [0, 1, 2, 3],
    kinds: ['format'],
  },
  {
    name: 'enter quote',
    body: p('Before') + p('First', true) + p('Second') + p('After'),
    target: 'Before\n\n> First\n>\n> Second\n\nAfter',
    remaining: ['Before', 'First', 'Second', 'After'],
    retained: [0, 1, 2, 3],
    kinds: ['format'],
  },
] as const;
for (const item of cases)
  test(`retained native quote ${item.name} keeps identities and reaches the exact target after save/reopen`, async () => {
    const input = await fixture(
      ('body' in item ? item.body : undefined) ??
        p('Before') + p('First', true) + p('Second', true) + p('After')
    );
    try {
      const main = input.pkg.parts.get(input.pkg.mainDocumentPart)!,
        body = bodyStoryRoot(main)!;
      if (body.kind === 'textValue') throw new Error('Missing body');
      const ids = body.children.filter((node) => node.kind === 'paragraph').map((node) => node.id),
        original = writeOoxmlPackage(input.pkg);
      const result = planTargetDocument(input, item.target);
      if (!result.success) throw new Error(result.error);
      expect([...new Set(result.plan.operations.map((edit) => edit.kind))].sort()).toEqual(
        [...item.kinds].sort()
      );
      const decoded = decodeEditPlan(serializeEditPlan(result.plan));
      expect(decoded.success).toBe(true);
      expect(result.pairs.map((pair) => pair.id)).toEqual(item.retained.map((index) => ids[index]));
      const direct = applyDirectPlan(input.editor, result.plan, sourceRevision);
      if (!direct.success) throw new Error(direct.error);
      const applied = input.editor.surface!.session.currentPackage(),
        nextMain = applied.parts.get(main.name)!,
        nextBody = bodyStoryRoot(nextMain)!;
      if (nextBody.kind === 'textValue') throw new Error('Missing body');
      expect(
        nextBody.children
          .filter((node) => node.kind === 'paragraph')
          .map((node) => paragraphTextOf(nextMain, node.id))
      ).toEqual([...item.remaining]);
      for (const index of item.retained)
        expect(findNode(nextMain, ids[index]!)?.kind).toBe('paragraph');
      expect(findNode(nextMain, ids[0]!)).toBe(findNode(main, ids[0]!));
      expect(findNode(nextMain, ids[3]!)).toBe(findNode(main, ids[3]!));
      expect(ast((await projectSourcePackage(applied)).text)).toEqual(ast(item.target));
      expect(canonicalOoxmlFingerprint(nextMain)).toContain('unchanged');
      const saved = writeOoxmlPackage(applied),
        reopened = readOoxmlPackage(saved);
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(ast((await projectSourcePackage(reopened.package)).text)).toEqual(ast(item.target));
      expect(writeOoxmlPackage(input.pkg)).toEqual(original);
    } finally {
      input.editor.destroy();
    }
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const scenario of ['suffix', 'delete-owner'] as const)
      test(`quote plan retains offline peer text through history and geometry (${scenario}); clients ${clientIds.join('/')}, receiver ${receiver}`, async () => {
        const seed = await fixture(p('Before') + p('First', true) + p('Second', true) + p('After'));
        const documentBytes = writeOoxmlPackage(seed.pkg);
        seed.editor.destroy();
        const pair = await createTwoEditorFixture({ document: documentBytes, clientIds });
        try {
          const [alice, bob] = pair.editors;
          const token = readEditorRevision(
            alice!,
            sourceRevision.documentId,
            sourceRevision.epoch
          )!;
          const working = await readWorkingView(alice!.surface!.session, token);
          if (!working.success) throw new Error(working.error);
          const initial = working.view.projection.text,
            ids = alice!.surface!.session.paragraphIds(),
            tail = ids[3]!;
          const target = 'Before\n\n> **First changed**\n>\n> New quote\n\nSecond\n\nAfter';
          const plan = planTargetDocument(
            {
              pkg: alice!.surface!.session.currentPackage(),
              revision: token,
              projection: working.view.projection,
            },
            target
          );
          if (!plan.success) throw new Error(plan.error);
          bob!.surface!.setSelection({
            anchor: { paragraphId: tail, offset: 5 },
            head: { paragraphId: tail, offset: 5 },
          });
          await pair.settle();
          pair.pause();
          const direct = applyDirectPlan(alice!, plan.plan, sourceRevision);
          if (!direct.success) throw new Error(direct.error);
          expect(
            bob!.surface!.applyAutomationOps(() => [
              scenario === 'suffix'
                ? { op: 'insertText', paragraphId: tail, offset: 5, text: ' Peer' }
                : { op: 'deleteBlock', blockId: ids[1]! },
            ]).rejected
          ).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          const expected =
            scenario === 'suffix' ? target + ' Peer' : 'Before\n\n> New quote\n\nSecond\n\nAfter';
          const undone = scenario === 'suffix' ? initial + ' Peer' : 'Before\n\n> Second\n\nAfter';
          for (const editor of pair.editors) {
            expect(
              ast((await projectSourcePackage(editor.surface!.session.currentPackage())).text)
            ).toEqual(ast(expected));
            for (const id of ids)
              if (scenario !== 'delete-owner' || id !== ids[1])
                expect(findNode(editor.surface!.session.part(), id)?.kind).toBe('paragraph');
              else expect(findNode(editor.surface!.session.part(), id)).toBeNull();
          }
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote).toBeDefined();
          expect(remote!.head.offset).toBe(bob!.surface!.state().selection!.head.offset);
          const position = { paragraphId: tail, offset: remote!.head.offset };
          const left = caretAt(alice!.surface!.layout(), position),
            right = caretAt(bob!.surface!.layout(), position);
          expect(left).not.toBeNull();
          expect(right).not.toBeNull();
          expect({ x: left!.x, y: left!.y }).toEqual({ x: right!.x, y: right!.y });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(
              ast((await projectSourcePackage(editor.surface!.session.currentPackage())).text)
            ).toEqual(ast(undone));
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(
              ast((await projectSourcePackage(editor.surface!.session.currentPackage())).text)
            ).toEqual(ast(expected));
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          for (const editor of pair.editors) {
            const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!saved.ok) throw new Error(saved.reason);
            expect(ast((await projectSourcePackage(saved.package)).text)).toEqual(ast(expected));
          }
        } finally {
          pair.destroy();
        }
      }, 30_000);
