import {
  findNode,
  prepareTrackedRunOrigins,
  replaceNode,
  sealTrackedRunOrigins,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { ancestorContentWriteRefusal } from '../blocks/ancestor-content-write';
import { buildBlockIndex } from '../blocks/block-index';
import { nativeCodeSource } from '../blocks/code-native-source';
import { ownedCodeSourceLock } from '../blocks/code-source-lock';
import { codeSourceOps } from '../blocks/code-source-ops';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

/** Release only exact preparation-owned locks within the caller's atomic transaction. */
export function codeSourceReviewLocks(pkg: OoxmlPackage, blockIds: readonly string[]) {
  const index = buildBlockIndex(pkg);
  const locks = blockIds.map((id) => {
    const block = index.byId.get(id);
    if (block?.payload.type !== 'code_block' || block.payload.version !== 2)
      throw Error('The native code owner is missing or unsupported.');
    const part = pkg.parts.get(block.partName)!;
    const control = findNode(part, block.controlId);
    const lock = control && ownedCodeSourceLock(control);
    const restriction = control && ancestorContentWriteRefusal(part, control.id);
    if (!control || !lock || restriction)
      throw Error(restriction ?? 'The native code lock has no exact preparation owner.');
    return { partName: part.name, controlId: control.id, lock };
  });
  const transform = (current: OoxmlPackage, release: boolean) => {
    let next = current;
    for (const { partName, controlId, lock } of locks) {
      const part = next.parts.get(partName)!;
      const control = findNode(part, controlId);
      if (!control || (release && ownedCodeSourceLock(control)?.id !== lock.id))
        throw Error('The native code lock changed before review.');
      const replaced = replaceNode(
        part,
        lock.id,
        release
          ? {
              ...lock,
              attributes: lock.attributes.map((attribute) =>
                attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'val'
                  ? { ...attribute, value: 'unlocked' }
                  : attribute
              ),
            }
          : lock
      );
      if (!replaced.ok) throw Error('The native code lock cannot retain its identity.');
      next = withPart(next, replaced.part);
    }
    return next;
  };
  return {
    before: (current: OoxmlPackage) => transform(current, true),
    after: (current: OoxmlPackage) => transform(current, false),
  };
}

/** Track source text and paragraph boundaries without replacing the code control or its lines. */
export function stageTrackedCodeSource(
  writer: TrackedStageWriter,
  blockId: string,
  before: string,
  after: string,
  attribution: TrackedStageAttribution
): void {
  const initial = writer.currentPackage();
  const block = buildBlockIndex(initial).byId.get(blockId);
  if (!block) throw Error('The native code block was deleted.');
  const initialPart = initial.parts.get(block.partName)!;
  const control = findNode(initialPart, block.controlId);
  const source = control && nativeCodeSource(initialPart, control);
  if (!source || source.value !== before)
    throw Error('The native code source changed or contains pending revisions.');
  const locks = codeSourceReviewLocks(initial, [blockId]);
  const ops = codeSourceOps(source, after, attribution);
  if (!ops.length) return;
  const touched = new Set(ops.flatMap((op) => ('paragraphId' in op ? [op.paragraphId] : [])));
  const scopes: { paragraphId: string; scope: string }[] = [];
  writer.edit(locks.before);
  for (let index = 0; index < source.paragraphs.length; index++) {
    const paragraph = source.paragraphs[index]!;
    if (!touched.has(paragraph.id)) continue;
    writer.edit((pkg) => {
      const prepared = prepareTrackedRunOrigins(pkg.parts.get(block.partName)!, paragraph.id, [
        { start: 0, end: source.lines[index]!.length },
      ]);
      if (prepared.scope) scopes.push({ paragraphId: paragraph.id, scope: prepared.scope });
      return withPart(pkg, prepared.part);
    });
  }
  for (const op of ops) writer.apply(block.partName, op);
  writer.edit((pkg) => {
    let part = pkg.parts.get(block.partName)!;
    for (const { paragraphId, scope } of scopes)
      part = sealTrackedRunOrigins(part, paragraphId, scope, initialPart);
    const current = nativeCodeSource(part, findNode(part, block.controlId)!, 'committed');
    if (!current) throw Error('The native source cannot retain its committed view.');
    const originalIds = new Set(source.paragraphs.map((paragraph) => paragraph.id));
    for (const paragraph of current.paragraphs)
      if (!originalIds.has(paragraph.id))
        for (const { scope } of scopes)
          part = sealTrackedRunOrigins(part, paragraph.id, scope, initialPart);
    const committed = nativeCodeSource(part, findNode(part, block.controlId)!, 'committed');
    const working = nativeCodeSource(part, findNode(part, block.controlId)!, 'working');
    if (committed?.value !== before || working?.value !== after)
      throw Error('The native code revision does not match its exact source views.');
    return withPart(pkg, part);
  });
  writer.edit(locks.after);
}
