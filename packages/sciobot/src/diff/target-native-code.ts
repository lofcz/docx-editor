import { findNode } from '@docx-editor.dev/core/store';
import { ancestorContentWriteRefusal } from '../blocks/ancestor-content-write';
import { buildBlockIndex } from '../blocks/block-index';
import { nativeCodeSource } from '../blocks/code-native-source';
import { ownedCodeSourceLock } from '../blocks/code-source-lock';
import type { BlockPayload } from '../blocks/payload-schema';
import { serializeCodeBlock } from '../projection/code-block';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { planAtomicBlock, type AtomicBlockSource } from './atomic-blocks';
import type { EditPlanResult } from './plan';
import { targetAstKey } from './target-key';

export const nativeCodeTargetPairable = (
  payload: BlockPayload,
  before: TargetAstNode,
  after: TargetAstNode
): boolean =>
  payload.type === 'code_block' &&
  payload.version === 2 &&
  before.type === 'code' &&
  after.type === 'code';

/** Route proven prepared code through its existing payload proposal and native source executor. */
export function planTargetNativeCode(
  source: AtomicBlockSource,
  payload: BlockPayload,
  before: TargetAstNode,
  after: TargetAstNode
): EditPlanResult {
  const fail = (error: string): EditPlanResult => ({
    success: false,
    code: 'unsupported-operation',
    error,
  });
  const owner = buildBlockIndex(source.pkg).byId.get(source.blockId);
  const part = owner && source.pkg.parts.get(owner.partName);
  const control = part && owner && findNode(part, owner.controlId);
  const code = part && control && nativeCodeSource(part, control);
  if (
    !owner ||
    !part ||
    !control ||
    !code ||
    !ownedCodeSourceLock(control) ||
    ancestorContentWriteRefusal(part, control.id) ||
    !nativeCodeTargetPairable(owner.payload, before, after)
  )
    return fail('The code target has no editable prepared native source owner.');
  if (
    payload.id !== owner.id ||
    payload.version !== 2 ||
    payload.type !== owner.payload.type ||
    payload.attributes.value !== code.value ||
    typeof after.value !== 'string' ||
    (after.lang !== null && after.lang !== undefined && typeof after.lang !== 'string') ||
    (after.meta !== null && after.meta !== undefined)
  )
    return fail('The code target contains stale or unsupported source syntax.');
  const matches = (value: BlockPayload, ast: TargetAstNode) => {
    const parsed = parseTargetMarkdown(serializeCodeBlock(value).text);
    return (
      parsed.success &&
      parsed.target.ast.children.length === 1 &&
      targetAstKey(parsed.target.ast.children[0]!) === targetAstKey(ast)
    );
  };
  if (!matches(owner.payload, before))
    return fail('The code source syntax differs from its native owner.');
  const target: BlockPayload = {
    ...owner.payload,
    attributes: { ...owner.payload.attributes, value: after.value, language: after.lang ?? '' },
  };
  if (!matches(target, after)) return fail('The code target cannot retain its exact syntax.');
  return planAtomicBlock(source, target);
}
