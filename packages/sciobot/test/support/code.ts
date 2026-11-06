import {
  readOoxmlPackage,
  removeNode,
  withPart,
  findNode,
  withCustomXmlNode,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { previewDirectPlan } from '../../src/ai/apply-direct';
import { withNativeCodeSources } from '../../src/blocks/code-preparation';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { ownedCodeSourceLock } from '../../src/blocks/code-source-lock';
import { readPayloadStore } from '../../src/blocks/payload-opc-codec';
import { encodeBlockPayload } from '../../src/blocks/payload-schema';
import { planTargetDocument } from '../../src/diff/plan-target';
import { targetDocumentInput } from './target-document';

export async function codeDocument(value = 'one\nmiddle\nlast', prepared = true) {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Tail</w:t></w:r></w:p>');
  const plan = planTargetDocument(input, `\`\`\`python\n${value}\n\`\`\`\n\nTail`, {
    allocateBlockId: () => 'code',
  });
  if (!plan.success) throw Error(plan.error);
  const applied = previewDirectPlan(input.pkg, plan.plan);
  if (!applied.success) throw Error(applied.error);
  if (prepared) return writeOoxmlPackage(withNativeCodeSources(applied.pkg));
  // New fragments prepare native source. A legacy fixture has neither its version nor owned lock.
  const block = buildBlockIndex(applied.pkg).byId.get('code')!;
  const part = applied.pkg.parts.get(block.partName)!;
  const control = findNode(part, block.controlId)!;
  const lock = ownedCodeSourceLock(control);
  const unlocked = lock ? removeNode(part, lock.id) : { ok: true as const, part };
  if (!unlocked.ok) throw Error('Cannot create legacy code fixture');
  const legacy = withPart(applied.pkg, unlocked.part);
  const store = readPayloadStore(legacy);
  const record = store.status === 'ready' && store.entries.get('code');
  if (store.status !== 'ready' || !record || record.decoded.status !== 'supported')
    throw Error('Missing code payload');
  return writeOoxmlPackage(
    withCustomXmlNode(legacy, store.part.partName, {
      id: 'scio_code',
      label: record.label,
      data: encodeBlockPayload({ ...record.decoded.value, version: 1 }),
    })
  );
}

export function openCode(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw Error(read.reason);
  return read.package;
}
