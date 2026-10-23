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
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../test/support/collaboration';
import { docxFixture } from '../test/support/document';
import { prepareNativeDocument } from '../src/native/prepare-document';
import { readEditorRevision } from '../src/projection/revision-cache';
import { readWorkingView } from '../src/review/working-view';
import { planTargetDocument } from '../src/diff/plan-target';
import { applyDirectPlan } from '../src/ai/apply-direct';

const prior = resolve(process.argv[2] ?? '');
assert.ok(process.argv[2], 'Pass the prior checkout path.');
assert.ok(
  !existsSync(join(prior, 'packages/pro/src/collaboration/document/paragraph-text-sources.ts')),
  'Use a reader without the new source implementation.'
);
assert.ok(
  !readFileSync(
    join(prior, 'packages/core/src/store/package/canonical-primitive-journal.ts'),
    'utf8'
  ).includes('readonly sourceText?'),
  'Use the earlier descriptor contract.'
);
const prepared = prepareNativeDocument(
  docxFixture(
    '<w:p><w:r><w:t xml:space="preserve">Keep. </w:t></w:r><w:hyperlink r:id="existing" w:tooltip="Read" w:history="1"><w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t>first and last</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> End.</w:t></w:r></w:p>',
    {
      relationships:
        '<Relationship Id="existing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://outer.test" TargetMode="External"/>',
    }
  )
);
if (!prepared.success) throw Error(prepared.error);
const pair = await createTwoEditorFixture({ document: prepared.document, clientIds: [41, 73] });
const directory = mkdtempSync(join(tmpdir(), 'docx-prior-text-reader-'));
try {
  const [alice, bob] = pair.editors;
  const paragraphId = alice!.surface!.session.paragraphIds()[0]!;
  const initial = alice!.surface!.session.part();
  const expected = applyTreeOp(initial, {
    op: 'insertText',
    paragraphId,
    offset: 'Keep. first '.length + 1,
    text: ' peer ',
  });
  if (!expected.ok) throw Error(expected.reason);
  const context = { documentId: 'prior-text-reader', epoch: 'native' };
  const revision = readEditorRevision(alice!, context.documentId, context.epoch)!;
  const working = await readWorkingView(alice!.surface!.session, revision);
  if (!working.success) throw Error(working.error);
  let id = 0;
  const planned = planTargetDocument(
    {
      revision,
      pkg: alice!.surface!.session.currentPackage(),
      projection: working.view.projection,
    },
    'Keep. [<code>first **$x$** and <source title="Paper" url="https://source.test" /> last</code>](https://outer.test "Read") End.',
    { allocateBlockId: () => `prior-reader-${++id}` }
  );
  if (!planned.success) throw Error(planned.error);
  const applied = applyDirectPlan(alice!, planned.plan, context);
  if (!applied.success) throw Error(applied.error);
  await pair.settle();
  const authored = alice!.surface!.session.part();
  const authoredPeer = applyTreeOp(authored, {
    op: 'insertText',
    paragraphId,
    offset: paragraphTextOf(authored, paragraphId)!.indexOf('and') + 1,
    text: ' peer ',
  });
  if (!authoredPeer.ok) throw Error(authoredPeer.reason);
  assert.equal(
    bob!.surface!.applyAutomationOps(
      () => [
        {
          op: 'insertText',
          paragraphId,
          offset: paragraphTextOf(bob!.surface!.session.part(), paragraphId)!.indexOf('and') + 1,
          text: ' peer ',
        },
      ],
      { kind: 'body' },
      [],
      false,
      true
    ).rejected,
    false
  );
  await pair.settle();
  assert.equal(alice!.exec({ type: 'undo' }).ok, true);
  await pair.settle();
  const expectedWord = canonicalOoxmlFingerprint(expected.part);
  for (const editor of pair.editors)
    assert.equal(canonicalOoxmlFingerprint(editor.surface!.session.part()), expectedWord);
  const offset = paragraphTextOf(expected.part, paragraphId)!.indexOf('peer') + 1;
  const legacy = applyTreeOp(expected.part, {
    op: 'insertText',
    paragraphId,
    offset,
    text: ' legacy ',
  });
  if (!legacy.ok) throw Error(legacy.reason);
  const legacyWord = canonicalOoxmlFingerprint(legacy.part);
  writeFileSync(
    join(directory, 'input.json'),
    JSON.stringify({
      document: Buffer.from(await alice!.save()).toString('base64'),
      room: Buffer.from(Y.encodeStateAsUpdate(pair.documents[0]!)).toString('base64'),
      paragraphId,
      offset,
    })
  );
  const child = `import { readFileSync, writeFileSync } from 'node:fs';
import * as Y from ${JSON.stringify(join(prior, 'packages/pro/node_modules/yjs/dist/yjs.mjs'))};
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from ${JSON.stringify(join(prior, 'packages/sciobot/test/support/collaboration.ts'))};
const input=JSON.parse(readFileSync(${JSON.stringify(join(directory, 'input.json'))},'utf8'));
const pair=await createTwoEditorFixture({ document: new Uint8Array(Buffer.from(input.document,'base64')), roomSnapshot: new Uint8Array(Buffer.from(input.room,'base64')), clientIds:[501,502] });
try {
 const snapshot=()=>({ words:pair.editors.map(e=>canonicalOoxmlFingerprint(e.surface.session.part())), update:Buffer.from(Y.encodeStateAsUpdate(pair.documents[0])).toString('base64') });
 const initial=snapshot();
 const result=pair.editors[0].surface.applyAutomationOps(()=>[{op:'insertText',paragraphId:input.paragraphId,offset:input.offset,text:' legacy '}],{kind:'body'},[],false,true);
 if(result.rejected)throw Error('Prior writer refused typing');
 await pair.settle();const edited=snapshot();
 if(!pair.editors[0].exec({type:'undo'}).ok)throw Error('Prior undo failed');
 await pair.settle();const undone=snapshot();
 if(!pair.editors[0].exec({type:'redo'}).ok)throw Error('Prior redo failed');
 await pair.settle();const redone=snapshot();
 writeFileSync(${JSON.stringify(join(directory, 'result.json'))},JSON.stringify({initial,states:[edited,undone,redone]}));
} finally {pair.destroy();}`;
  writeFileSync(join(directory, 'prior.ts'), child);
  const ran = spawnSync(
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
  if (ran.error) throw ran.error;
  assert.equal(ran.status, 0);
  const output = JSON.parse(readFileSync(join(directory, 'result.json'), 'utf8'));
  for (const word of output.initial.words) assert.equal(word, expectedWord);
  for (const [index, state] of output.states.entries()) {
    const word = index === 1 ? expectedWord : legacyWord;
    for (const priorWord of state.words) assert.equal(priorWord, word);
    Y.applyUpdate(
      pair.documents[0]!,
      new Uint8Array(Buffer.from(state.update, 'base64')),
      'prior-reader'
    );
    await pair.settle();
    for (const editor of pair.editors)
      assert.equal(canonicalOoxmlFingerprint(editor.surface!.session.part()), word);
  }
  assert.equal(JSON.stringify(pair.documents[0]!.toJSON()).includes('"sourceText"'), false);
  const redo = applyTreeOp(authoredPeer.part, {
    op: 'insertText',
    paragraphId,
    offset: paragraphTextOf(authoredPeer.part, paragraphId)!.indexOf('peer') + 1,
    text: ' legacy ',
  });
  if (!redo.ok) throw Error(redo.reason);
  assert.equal(alice!.exec({ type: 'redo' }).ok, true);
  await pair.settle();
  for (const editor of pair.editors)
    assert.equal(
      canonicalOoxmlFingerprint(editor.surface!.session.part()),
      canonicalOoxmlFingerprint(redo.part)
    );
  console.log(
    'Prior reader/writer compatibility passes 16 exact native snapshots; shared state contains no sourceText field.'
  );
} finally {
  pair.destroy();
  rmSync(directory, { recursive: true, force: true });
}
