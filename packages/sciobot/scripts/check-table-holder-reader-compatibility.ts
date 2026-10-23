import { reviewModule } from '../../pro/src/index.ts';
// Test-only compatibility transport uses the EigenPal Pro License.
import { strict as assert } from 'node:assert';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import * as Y from 'yjs';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../test/support/collaboration';
import { docxFixture } from '../test/support/document';

const prior = resolve(process.argv[2] ?? '');
assert.ok(process.argv[2], 'Pass the prior checkout path.');
const previous = join(
  prior,
  'packages/core/src/store/store/inserted-paragraph-property-presence.ts'
);
assert.ok(
  existsSync(previous) &&
    !readFileSync(previous, 'utf8').includes('withInsertedTablePropertyPresence'),
  'Use a reader without table property presence support.'
);
const directory = mkdtempSync(join(tmpdir(), 'docx-prior-table-holder-'));
let snapshots = 0;
function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
function elements(part: OoxmlPart, kind: OoxmlNode['kind']): OoxmlElement[] {
  const result: OoxmlElement[] = [];
  const walk = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === kind) result.push(node);
    node.children.forEach(walk);
  };
  walk(part.root);
  return result;
}
try {
  for (const topology of ['row', 'column'] as const) {
    const pair = await createTwoEditorFixture({
      document: docxFixture(
        '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Old A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Old B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
      ),
      clientIds: [41, 73],
      editor: { modules: [reviewModule()] },
    });
    try {
      const first = pair.editors[0]!;
      const initial = first.surface!.session.part();
      const initialIds = new Set(first.surface!.session.paragraphIds());
      const oldParagraph = elements(initial, 'paragraph').find(
        (node) => paragraphTextOf(initial, node.id) === 'Old A'
      )!.id;
      const table = elements(initial, 'table')[0]!,
        row = elements(initial, 'tableRow')[0]!;
      const grid = table.children.find(
        (node) => node.kind !== 'textValue' && node.localName === 'tblGrid'
      ) as OoxmlElement;
      const revision = { author: 'Current', date: '2026-10-23T09:00:00Z' };
      const op: TreeDocOp =
        topology === 'row'
          ? { op: 'insertTableRow', tableId: table.id, rowId: row.id, where: 'below', revision }
          : {
              op: 'insertTableColumn',
              tableId: table.id,
              gridColumnId: grid.children[0]!.id,
              where: 'left',
              revision,
            };
      const birth = first.surface!.applyAutomationOps(
        () => [op],
        { kind: 'body' },
        [],
        false,
        true
      );
      assert.equal(birth.rejected, false, `${topology}: ${JSON.stringify(birth)}`);
      await pair.settle();
      const bornParagraph = first
        .surface!.session.paragraphIds()
        .find((id) => !initialIds.has(id))!;
      assert.ok(bornParagraph);
      const pending = first.surface!.session.part();
      for (const editor of pair.editors) {
        assert.equal(
          canonicalOoxmlFingerprint(editor.surface!.session.part()),
          canonicalOoxmlFingerprint(pending)
        );
        snapshots++;
      }
      let phase = 0;
      const olderWrites = async (paragraphId: string, value: string) => {
        const before = first.surface!.session.part();
        const offset = paragraphTextOf(before, paragraphId)!.length;
        const edited = apply(before, { op: 'insertText', paragraphId, offset, text: value });
        const expected = [edited, before, edited].map(canonicalOoxmlFingerprint);
        const inputPath = join(directory, 'input.json'),
          outputPath = join(directory, 'output.json');
        const clients = [501 + phase * 2, 502 + phase * 2];
        phase += 1;
        writeFileSync(
          inputPath,
          JSON.stringify({
            document: Buffer.from(await first.save()).toString('base64'),
            room: Buffer.from(Y.encodeStateAsUpdate(pair.documents[0]!)).toString('base64'),
            paragraphId,
            offset,
            value,
            clients,
          })
        );
        const child = `import { readFileSync, writeFileSync } from 'node:fs';
import * as Y from ${JSON.stringify(join(prior, 'packages/pro/node_modules/yjs/dist/yjs.mjs'))};
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from ${JSON.stringify(join(prior, 'packages/sciobot/test/support/collaboration.ts'))};
const input=JSON.parse(readFileSync(${JSON.stringify(inputPath)},'utf8'));
const pair=await createTwoEditorFixture({document:new Uint8Array(Buffer.from(input.document,'base64')),roomSnapshot:new Uint8Array(Buffer.from(input.room,'base64')),clientIds:input.clients});
try {
 const snapshot=()=>({words:pair.editors.map(e=>canonicalOoxmlFingerprint(e.surface.session.part())),update:Buffer.from(Y.encodeStateAsUpdate(pair.documents[0])).toString('base64')});
 const initial=snapshot();
 const result=pair.editors[0].surface.applyAutomationOps(()=>[{op:'insertText',paragraphId:input.paragraphId,offset:input.offset,text:input.value}],{kind:'body'},[],false,true);
 if(result.rejected)throw Error('Prior writer refused table typing');
 await pair.settle();const edited=snapshot();
 if(!pair.editors[0].exec({type:'undo'}).ok)throw Error('Prior undo failed');
 await pair.settle();const undone=snapshot();
 if(!pair.editors[0].exec({type:'redo'}).ok)throw Error('Prior redo failed');
 await pair.settle();const redone=snapshot();
 writeFileSync(${JSON.stringify(outputPath)},JSON.stringify({initial,states:[edited,undone,redone]}));
}finally{pair.destroy();}`;
        writeFileSync(join(directory, 'prior.ts'), child);
        const result = spawnSync(
          'bun',
          [
            '--preload',
            join(prior, 'scripts/test/preload-dom.ts'),
            '--tsconfig-override',
            join(prior, 'tsconfig.json'),
            join(directory, 'prior.ts'),
          ],
          { cwd: prior, stdio: 'inherit' }
        );
        if (result.error) throw result.error;
        assert.equal(result.status, 0);
        const output = JSON.parse(readFileSync(outputPath, 'utf8'));
        for (const word of output.initial.words) {
          assert.equal(word, canonicalOoxmlFingerprint(before));
          snapshots++;
        }
        for (const [index, state] of output.states.entries()) {
          for (const word of state.words) {
            assert.equal(word, expected[index]);
            snapshots++;
          }
          Y.applyUpdate(
            pair.documents[0]!,
            new Uint8Array(Buffer.from(state.update, 'base64')),
            'prior-table-reader'
          );
          await pair.settle();
          for (const editor of pair.editors) {
            assert.equal(
              canonicalOoxmlFingerprint(editor.surface!.session.part()),
              expected[index]
            );
            snapshots++;
          }
        }
      };
      await olderWrites(oldParagraph, ' prior ');
      const beforeAcceptance = first.surface!.session.part();
      const accepted = apply(beforeAcceptance, { op: 'acceptAllRevisions' });
      assert.equal(
        first.surface!.applyAutomationOps(
          () => [{ op: 'acceptAllRevisions' }],
          { kind: 'body' },
          [],
          false,
          true
        ).rejected,
        false
      );
      await pair.settle();
      for (const editor of pair.editors) {
        assert.equal(
          canonicalOoxmlFingerprint(editor.surface!.session.part()),
          canonicalOoxmlFingerprint(accepted)
        );
        snapshots++;
      }
      await olderWrites(bornParagraph, ' prior cell ');
      const acceptedPeer = first.surface!.session.part();
      const pendingPeer = apply(beforeAcceptance, {
        op: 'insertText',
        paragraphId: bornParagraph,
        offset: 0,
        text: ' prior cell ',
      });
      assert.equal(first.exec({ type: 'undo' }).ok, true);
      await pair.settle();
      for (const editor of pair.editors) {
        assert.equal(
          canonicalOoxmlFingerprint(editor.surface!.session.part()),
          canonicalOoxmlFingerprint(pendingPeer)
        );
        snapshots++;
      }
      assert.equal(first.exec({ type: 'redo' }).ok, true);
      await pair.settle();
      for (const editor of pair.editors) {
        assert.equal(
          canonicalOoxmlFingerprint(editor.surface!.session.part()),
          canonicalOoxmlFingerprint(acceptedPeer)
        );
        snapshots++;
      }
    } finally {
      pair.destroy();
    }
  }
  console.log(
    `Prior table reader/writer compatibility passes ${snapshots} exact native snapshots before and after current-writer acceptance.`
  );
} finally {
  rmSync(directory, { recursive: true, force: true });
}
