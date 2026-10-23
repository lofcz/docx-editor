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
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../test/support/collaboration';
import { docxFixture } from '../test/support/document';
const prior = resolve(process.argv[2] ?? '');
assert.ok(process.argv[2], 'Pass the prior checkout path.');
const validator = join(prior, 'packages/core/src/store/store/tree-op-validate.ts');
assert.ok(
  existsSync(validator) &&
    !existsSync(join(prior, 'packages/core/src/store/store/named-text-owner.ts')),
  'Use a reader before hyperlink text ownership support.'
);
function nodes(node: OoxmlNode): OoxmlNode[] {
  return [node, ...(node.kind === 'textValue' ? [] : node.children.flatMap(nodes))];
}
function apply(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
const directory = mkdtempSync(join(tmpdir(), 'docx-prior-hyperlink-text-'));
let snapshots = 0;
try {
  const pair = await createTwoEditorFixture({
    document: docxFixture(
      '<w:p><w:r><w:t>Keep. </w:t></w:r><w:hyperlink r:id="old" w:tooltip="Read" w:history="1" w:tgtFrame="_blank"><w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t>label</w:t></w:r></w:hyperlink><w:r><w:t> End.</w:t></w:r></w:p>',
      {
        relationships:
          '<Relationship Id="old" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://outer.test" TargetMode="External"/>',
      }
    ),
    clientIds: [41, 73],
  });
  try {
    const first = pair.editors[0]!,
      original = first.surface!.session.part(),
      paragraphId = first.surface!.session.paragraphIds()[0]!,
      link = nodes(original.root).find((n) => n.kind === 'hyperlink')!;
    const insertion: TreeDocOp = {
      op: 'insertText',
      paragraphId,
      offset: 11,
      text: ' ',
      inside: link.id,
    };
    const authored = apply(original, insertion);
    assert.equal(
      first.surface!.applyAutomationOps(() => [insertion], { kind: 'body' }, [], false, true)
        .rejected,
      false
    );
    await pair.settle();
    const compare = (expected: OoxmlPart) => {
      for (const editor of pair.editors) {
        assert.equal(
          canonicalOoxmlFingerprint(editor.surface!.session.part()),
          canonicalOoxmlFingerprint(expected)
        );
        snapshots++;
      }
    };
    compare(authored);
    const inputPath = join(directory, 'input.json'),
      outputPath = join(directory, 'output.json');
    writeFileSync(
      inputPath,
      JSON.stringify({
        document: Buffer.from(await first.save()).toString('base64'),
        room: Buffer.from(Y.encodeStateAsUpdate(pair.documents[0]!)).toString('base64'),
        paragraphId,
        offset: 12,
      })
    );
    const child = `import {readFileSync,writeFileSync} from 'node:fs';
import * as Y from ${JSON.stringify(join(prior, 'packages/pro/node_modules/yjs/dist/yjs.mjs'))};
import {canonicalOoxmlFingerprint} from '@docx-editor.dev/core/store';
import {createTwoEditorFixture} from ${JSON.stringify(join(prior, 'packages/sciobot/test/support/collaboration.ts'))};
const input=JSON.parse(readFileSync(${JSON.stringify(inputPath)},'utf8'));
const pair=await createTwoEditorFixture({document:new Uint8Array(Buffer.from(input.document,'base64')),roomSnapshot:new Uint8Array(Buffer.from(input.room,'base64')),clientIds:[501,502]});
try{
const snapshot=()=>({words:pair.editors.map(e=>canonicalOoxmlFingerprint(e.surface.session.part())),update:Buffer.from(Y.encodeStateAsUpdate(pair.documents[0])).toString('base64')});
const initial=snapshot();
const write=pair.editors[0].surface.applyAutomationOps(()=>[{op:'insertText',paragraphId:input.paragraphId,offset:input.offset,text:' prior'}],{kind:'body'},[],false,true);if(write.rejected)throw Error('Prior writer refused outside-link typing');
await pair.settle();const edited=snapshot();
if(!pair.editors[0].exec({type:'undo'}).ok)throw Error('Prior undo failed');await pair.settle();const undone=snapshot();
if(!pair.editors[0].exec({type:'redo'}).ok)throw Error('Prior redo failed');await pair.settle();const redone=snapshot();
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
      assert.equal(word, canonicalOoxmlFingerprint(authored));
      snapshots++;
    }
    const edited = apply(authored, { op: 'insertText', paragraphId, offset: 12, text: ' prior' });
    for (const [i, state] of output.states.entries()) {
      const expected = [edited, authored, edited][i]!;
      for (const word of state.words) {
        assert.equal(word, canonicalOoxmlFingerprint(expected));
        snapshots++;
      }
      Y.applyUpdate(
        pair.documents[0]!,
        new Uint8Array(Buffer.from(state.update, 'base64')),
        'prior-hyperlink-text-reader'
      );
      await pair.settle();
      compare(expected);
    }
    assert.equal(first.exec({ type: 'undo' }).ok, true);
    await pair.settle();
    compare(apply(original, { op: 'insertText', paragraphId, offset: 11, text: ' prior' }));
    assert.equal(first.exec({ type: 'redo' }).ok, true);
    await pair.settle();
    compare(edited);
  } finally {
    pair.destroy();
  }
  console.log(
    `Prior hyperlink text reader/writer compatibility passes ${snapshots} exact native snapshots through both histories.`
  );
} finally {
  rmSync(directory, { recursive: true, force: true });
}
