import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  findNode,
  replaceNode,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { buildBlockIndex } from './block-index';
import { nativeCodeSource } from './code-native-source';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { DEFAULT_PAYLOAD_LIMITS, encodeBlockPayload } from './payload-schema';
import type { BlockPayload } from './payload-schema';
import { fieldChanges } from './payload-fields';
import { ownedCodeSourceLock } from './code-source-lock';
import { codeSourceOps } from './code-source-ops';

type CodeEditResult = { readonly success: true; readonly revision: number } | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

function codeTarget(editor: DocxEditorInstance, blockId: string, metadataOnly = false) {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Code changes require edit mode.');
  const pkg = surface.session.currentPackage();
  const block = buildBlockIndex(pkg).byId.get(blockId);
  if (!block)
    return fail('stale-revision', 'The code block was deleted. Your source remains available.');
  if (block.payload.version !== 2 || block.payload.type !== 'code_block')
    return fail(
      'unsupported-operation',
      'Prepare plain code source before editing. Imported structures remain unchanged.'
    );
  const part = pkg.parts.get(block.partName)!;
  const control = findNode(part, block.controlId);
  const source = control && nativeCodeSource(part, control, metadataOnly ? 'committed' : 'plain');
  if (!control || !source)
    return fail('unsupported-operation', 'This code block contains unsupported native content.');
  const properties = contentControlPropertiesOf(control);
  const ancestor = ancestorContentWriteRefusal(part, control.id);
  if (ancestor || properties.dataBinding)
    return fail('unsupported-operation', ancestor ?? 'The code block is bound to custom XML.');
  const lockNode = ownedCodeSourceLock(control);
  if (!lockNode)
    return fail('unsupported-operation', 'The code source lock is not owned by this editor.');
  return { success: true as const, surface, pkg, block, part, control, source, lockNode };
}

/** Change language without taking ownership of source or unrelated payload fields. */
export function updateCodeLanguage(
  editor: DocxEditorInstance,
  request: {
    readonly blockId: string;
    readonly language: string;
  }
): CodeEditResult {
  const target = codeTarget(editor, request.blockId, true);
  if (!target.success) return target;
  if (typeof request.language !== 'string' || request.language.length > 128)
    return fail('invalid-input', 'The code language is invalid.');
  const { surface, block } = target;
  if (block.payload.attributes.language === request.language)
    return { success: true, revision: surface.session.packageRevision() };
  const transform = (pkg: OoxmlPackage) => {
    const live = buildBlockIndex(pkg).byId.get(block.id);
    if (!live) throw Error('The code block was deleted.');
    const store = readPayloadStore(pkg);
    if (store.status !== 'ready') throw Error('The code payload is missing.');
    const written = withPayloadRecord(
      pkg,
      { ...live.payload, attributes: { ...live.payload.attributes, language: request.language } },
      store.entries.get(block.id)!.label
    );
    if (!written.ok) throw Error(written.reason);
    return written.pkg;
  };
  try {
    transform(target.pkg);
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true);
    return result.rejected || !result.committed
      ? fail('apply-failed', String(result.reason))
      : { success: true, revision: surface.session.packageRevision() };
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : String(error));
  }
}

/** Commit source through ordinary native text and paragraph operations in one history unit. */
export function updateCodeSource(
  editor: DocxEditorInstance,
  request: {
    readonly blockId: string;
    readonly expectedSource: string;
    readonly value: string;
  }
): CodeEditResult {
  return commitCodeSource(editor, request);
}

/** Existing complete-payload callers share the source transaction instead of replacing its paragraphs. */
export function updateCodePayload(
  editor: DocxEditorInstance,
  original: BlockPayload,
  next: BlockPayload
): CodeEditResult {
  return commitCodeSource(
    editor,
    {
      blockId: original.id,
      expectedSource: original.attributes.value as string,
      value: next.attributes.value as string,
    },
    next.attributes
  );
}

function commitCodeSource(
  editor: DocxEditorInstance,
  request: {
    readonly blockId: string;
    readonly expectedSource: string;
    readonly value: string;
  },
  attributes?: BlockPayload['attributes']
): CodeEditResult {
  const target = codeTarget(editor, request.blockId, request.value === request.expectedSource);
  if (!target.success) return target;
  const { surface, block, source, lockNode } = target;
  if (
    typeof request.value !== 'string' ||
    typeof request.expectedSource !== 'string' ||
    request.value.length > DEFAULT_PAYLOAD_LIMITS.maxBytes ||
    /\r/.test(request.value) ||
    request.value.split('\n').length > 10000
  )
    return fail(
      'invalid-input',
      'Code source needs supported line endings and at most 10000 lines.'
    );
  if (request.expectedSource !== source.value)
    return fail('stale-revision', 'The code source changed. Your source remains available.');
  if (request.value === source.value && !attributes)
    return { success: true, revision: surface.session.packageRevision() };
  try {
    const payload = {
      ...block.payload,
      attributes: { ...(attributes ?? block.payload.attributes), value: request.value },
    };
    encodeBlockPayload(payload);
    if (!fieldChanges(block.payload, payload).length)
      return { success: true, revision: surface.session.packageRevision() };
    const ops = codeSourceOps(source, request.value);
    const locked = ops.length > 0;
    const lockPart = (part: OoxmlPart, unlock: boolean) => {
      if (!locked) return part;
      if (!lockNode || lockNode.kind !== 'generic') throw Error('The code lock is invalid.');
      // This transaction owns the app lock, but never its identity. Recreating it would
      // give concurrent writers different property nodes and corrupt the native control.
      const changed = replaceNode(
        part,
        lockNode.id,
        unlock
          ? {
              ...lockNode,
              attributes: lockNode.attributes.map((attribute) =>
                attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'val'
                  ? { ...attribute, value: 'unlocked' }
                  : attribute
              ),
            }
          : lockNode
      );
      if (!changed.ok) throw Error('The code lock cannot be retained.');
      return changed.part;
    };
    const run = (readPart: () => OoxmlPart, apply: (op: TreeDocOp) => void) => {
      for (const op of ops) apply(op);
      const part = readPart();
      const changed = nativeCodeSource(part, findNode(part, block.controlId)!);
      if (changed?.value !== request.value)
        throw Error('The native code source did not match the requested edit.');
    };
    const writeMetadata = (pkg: OoxmlPackage) => {
      if (!attributes) return pkg;
      const store = readPayloadStore(pkg);
      if (store.status !== 'ready') throw Error('The code payload is missing.');
      const stored = withPayloadRecord(pkg, payload, store.entries.get(block.id)!.label);
      if (!stored.ok) throw Error(stored.reason);
      return stored.pkg;
    };
    const transform = (pkg: OoxmlPackage) => {
      let part = lockPart(pkg.parts.get(block.partName)!, true);
      run(
        () => part,
        (op) => {
          const result = applyTreeOp(part, op);
          if (!result.ok) throw Error(`The native code edit was refused: ${result.reason}.`);
          part = result.part;
        }
      );
      return writeMetadata(withPart(pkg, lockPart(part, false)));
    };
    transform(target.pkg);
    const result = surface.applyAutomationOps(
      () => [],
      { kind: 'body' },
      undefined,
      false,
      true,
      'before',
      (ctx) => {
        const lock = (unlock: boolean) => {
          if (
            !ctx.applyPackage((pkg) =>
              withPart(pkg, lockPart(pkg.parts.get(block.partName)!, unlock))
            )
          )
            throw Error('The code lock transaction was refused.');
        };
        lock(true);
        run(
          () => ctx.currentPackage().parts.get(block.partName)!,
          (op) => {
            if (!ctx.applyTo(block.partName, op))
              throw Error('The native code transaction refused an operation.');
          }
        );
        lock(false);
        if (!ctx.applyPackage(writeMetadata))
          throw Error('The code metadata transaction was refused.');
      }
    );
    return result.rejected || !result.committed
      ? fail('apply-failed', String(result.reason))
      : { success: true, revision: surface.session.packageRevision() };
  } catch (error) {
    return fail('apply-failed', error instanceof Error ? error.message : String(error));
  }
}
