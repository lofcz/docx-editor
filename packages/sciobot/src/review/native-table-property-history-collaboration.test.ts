import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { cellReviewDocument, cellScopeFor, CELL_CONTEXT } from '../../test/support/table-review';
import { groups } from '../../test/support/review-groups';
import { planTableMerge } from '../diff/table-merges';
import { readEditorRevision } from '../projection/revision-cache';
import { trackTableMergePlan } from './track-table-merges';
import { acceptReviewGroup } from './accept-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const revision = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
function word(pkg: OoxmlPackage) {
  const before = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toBe(before);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
function resolvedGeometry(editor: DocxEditorInstance, spanning: boolean) {
  const all = (node: OoxmlNode): readonly OoxmlNode[] => [
    node,
    ...(node.kind === 'textValue' ? [] : node.children.flatMap(all)),
  ];
  const table = all(editor.surface!.session.part().root).find((node) => node.kind === 'table')!;
  const row = table.children.find((node) => node.kind === 'tableRow')!;
  const owner = row.children.find((node) => node.kind === 'tableCell')!;
  const properties = owner.children.filter((node) => node.localName === 'tcPr');
  expect(properties).toHaveLength(1);
  const spans = all(properties[0]!).filter(
    (node) => node.kind !== 'textValue' && node.localName === 'gridSpan'
  );

  expect(spans).toHaveLength(spanning ? 1 : 0);
  if (!spanning) return;
  const span = spans[0]!;
  expect(
    span.kind !== 'textValue' &&
      span.attributes.find((attribute) => attribute.localName === 'val')?.value
  ).toBe('2');
}
function imported(hasProperties: boolean, keepColumns: boolean, nested = false) {
  const p = (text: string) => '<w:p><w:r><w:t>' + text + '</w:t></w:r></w:p>';
  const removed = nested
    ? '<w:sdt><w:sdtPr/><w:sdtContent>' + p('B') + '</w:sdtContent></w:sdt>'
    : p('B');
  return docxFixture(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>' +
      (hasProperties
        ? '<w:tcPr><w:shd w:fill="EEEEEE"/><w:customKeep w:val="retain"/></w:tcPr>'
        : '') +
      p('A') +
      '</w:tc><w:tc><w:tcPr><w:shd w:fill="BBBBBB"/>' +
      '<x:keep xmlns:x="urn:r51:source" x:val="original"><x:nested/></x:keep>' +
      '<w:cellDel w:id="9" w:author="Teacher"/></w:tcPr>' +
      removed +
      '</w:tc></w:tr>' +
      (keepColumns
        ? '<w:tr><w:tc>' + p('One') + '</w:tc><w:tc>' + p('Two') + '</w:tc></w:tr>'
        : '') +
      '</w:tbl>' +
      p('Tail')
  );
}
function importedVertical() {
  const cell = (text: string, merge: string) =>
    '<w:tr><w:tc><w:tcPr><w:shd w:fill="EEEEEE"/><w:customKeep w:val="retain"/>' +
    '<w:cellMerge w:id="9" w:author="Teacher" w:vMerge="' +
    merge +
    '"/>' +
    '</w:tcPr><w:p><w:r><w:t>' +
    text +
    '</w:t></w:r></w:p></w:tc></w:tr>';
  return docxFixture(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      cell('A', 'rest') +
      cell('B', 'cont') +
      '</w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
  );
}
for (const mode of [
  'tracked-merge',
  'tracked-vertical-merge',
  'imported-vertical-merge',
  'imported-properties',
  'imported-no-properties',
  'imported-merge-properties',
  'imported-merge-no-properties',
  'imported-merge-nested-no-properties',
] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      for (const first of [0, 1] as const)
        test(
          'R51 concurrent single-group ' +
            mode +
            ' keeps one exact Word property set through both actors history; ' +
            clients +
            '/' +
            receiver +
            '/' +
            first,
          async () => {
            const pair = await createTwoEditorFixture({
              document: mode.startsWith('tracked-')
                ? cellReviewDocument()
                : mode === 'imported-vertical-merge'
                  ? importedVertical()
                  : imported(
                      !mode.endsWith('no-properties'),
                      mode.startsWith('imported-merge-'),
                      mode.includes('nested')
                    ),
              ownedAttributes: true,
              clientIds: clients,
              editor: options,
            });
            try {
              const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
              if (mode.startsWith('tracked-')) {
                const scope = cellScopeFor(alice),
                  plan = planTableMerge(
                    {
                      part: alice.surface!.session.part(),
                      revision: revision(alice),
                      tableId: scope.tableId,
                    },
                    {
                      kind: 'merge',
                      cellIds: [
                        cellScopeFor(alice, 'A').cellId,
                        cellScopeFor(alice, mode === 'tracked-vertical-merge' ? 'One' : 'B').cellId,
                      ],
                    }
                  );
                if (!plan.success) throw new Error(plan.error);
                expect(trackTableMergePlan(alice, plan.plan, CELL_CONTEXT).success).toBe(true);
                await pair.settle();
              }
              const peer = bob
                .surface!.session.paragraphIds()
                .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Tail')!;
              bob.surface!.setSelection({
                anchor: { paragraphId: peer, offset: 4 },
                head: { paragraphId: peer, offset: 4 },
              });
              bob.surface!.insertPlainText(' peer');
              bob.surface!.flushPendingInput();
              await pair.settle();
              const pending = groups(alice);
              if (mode.endsWith('vertical-merge'))
                expect(
                  collectRevisionSites(alice.surface!.session.part()).filter(
                    (site) => site.node.localName === 'cellMerge'
                  )
                ).toHaveLength(2);
              expect(pending).toHaveLength(1);
              expect(pending[0]!.readOnly).toBe(false);
              const pkg = alice.surface!.session.currentPackage(),
                native = applyTreeOp(alice.surface!.session.part(), { op: 'acceptAllRevisions' });
              if (!native.ok) throw new Error(native.reason);
              const expected = word(withPart(pkg, native.part)),
                original = word(pkg),
                count = collectRevisionSites(alice.surface!.session.part()).length,
                id = pending[0]!.id;
              pair.pause();
              for (const editor of pair.editors)
                expect(acceptReviewGroup(editor, id, revision(editor), CELL_CONTEXT)).toMatchObject(
                  {
                    success: true,
                  }
                );
              pair.reconnect(receiver);
              await pair.settle();
              const check = (resolved: boolean) => {
                for (const editor of pair.editors) {
                  expect(groups(editor).map((group) => group.id)).toEqual(resolved ? [] : [id]);
                  expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(
                    resolved ? 0 : count
                  );
                  if (resolved)
                    resolvedGeometry(
                      editor,
                      mode === 'tracked-merge' || mode.startsWith('imported-merge-')
                    );
                  expect(word(editor.surface!.session.currentPackage())).toBe(
                    resolved ? expected : original
                  );
                  expect(paragraphTextOf(editor.surface!.session.part(), peer)).toBe('Tail peer');
                }
                const remote = pair.rooms[0]!.session.remoteSelections().find(
                  (row) => row.actorId === 'bob'
                );
                expect(remote).toBeDefined();
                expect(
                  caretAt(alice.surface!.layout(), {
                    paragraphId: remote!.head.nodeId,
                    offset: remote!.head.offset,
                  })
                ).not.toBeNull();
              };
              check(true);
              expect(pair.editors[first]!.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(true);
              expect(pair.editors[1 - first]!.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(false);
              expect(pair.editors[first]!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(true);
              expect(pair.editors[1 - first]!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(true);
              expect(pair.editors[first]!.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(true);
              expect(pair.editors[1 - first]!.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              check(false);
              expect(pair.editors[first]!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(true);
              expect(pair.editors[1 - first]!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              check(true);
              const digests = await pair.savedDigests();
              expect(digests[0]).toEqual(digests[1]);
              const reopened = await createTwoEditorFixture({
                document: new Uint8Array(await alice.save()),
                roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
                ownedAttributes: true,
                editor: options,
              });
              try {
                for (const editor of reopened.editors) {
                  expect(groups(editor)).toEqual([]);
                  expect(word(editor.surface!.session.currentPackage())).toBe(expected);
                }
              } finally {
                reopened.destroy();
              }
            } finally {
              pair.destroy();
            }
          },
          30000
        );

for (const nested of [false, true])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(
        'R51 cell geometry retains unseen peer text through removal, history, and reopening; ' +
          [nested, clients, receiver],
        async () => {
          const pair = await createTwoEditorFixture({
            document: imported(false, true, nested),
            ownedAttributes: true,
            clientIds: clients,
            editor: options,
          });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
              paragraphId = bob
                .surface!.session.paragraphIds()
                .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'B')!,
              id = groups(alice)[0]!.id;
            pair.pause();
            expect(acceptReviewGroup(alice, id, revision(alice), CELL_CONTEXT)).toMatchObject({
              success: true,
            });
            const acceptedPackage = alice.surface!.session.currentPackage(),
              accepted = word(acceptedPackage),
              future = applyTreeOp(alice.surface!.session.part(), {
                op: 'insertText',
                paragraphId,
                offset: 0,
                text: ' peer',
              });
            if (!future.ok) throw new Error(future.reason);
            const expected = word(withPart(acceptedPackage, future.part));
            bob.surface!.setSelection({
              anchor: { paragraphId, offset: 1 },
              head: { paragraphId, offset: 1 },
            });
            bob.surface!.insertPlainText(' peer');
            bob.surface!.flushPendingInput();
            const pending = word(bob.surface!.session.currentPackage());
            pair.reconnect(receiver);
            await pair.settle();
            const check = (expectedWord: string, pendingGroup: boolean) => {
              for (const editor of pair.editors) {
                expect(word(editor.surface!.session.currentPackage())).toBe(expectedWord);
                expect(groups(editor).map((group) => group.id)).toEqual(pendingGroup ? [id] : []);
              }
            };
            check(expected, false);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            check(pending, true);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            check(expected, false);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            check(accepted, false);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            check(expected, false);
            const reopened = await createTwoEditorFixture({
              document: new Uint8Array(await alice.save()),
              roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
              ownedAttributes: true,
              editor: options,
            });
            try {
              for (const editor of reopened.editors)
                expect(word(editor.surface!.session.currentPackage())).toBe(expected);
            } finally {
              reopened.destroy();
            }
          } finally {
            pair.destroy();
          }
        },
        30000
      );
