import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  contentControlPropertiesOf,
  contentControlsIn,
  createNodeIdAllocator,
  findNode,
  readOoxmlPart,
  replaceChildren,
  usedParaIds,
  w14RootPrefix,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { JsonValue } from '../data/json';
import { runXml } from '../diff/fragment-seeds';
import { withInlineCodeStyle } from '../native/inline-code';
import { activeMdxRegistry } from '../registry/active';
import { MdxRefusal } from '../registry/mdx-fields';
import { buildBlockIndex } from './block-index';
import { withFreshIds } from './fresh-ids';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import type { BlockPayload } from './payload-schema';

export interface AtomicBlockUpdate {
  readonly blockId: string;
  /** The complete new attributes. The payload type and ID stay. */
  readonly attributes: Readonly<Record<string, JsonValue>>;
  readonly expectedRevision?: number;
}

export type AtomicBlockUpdateResult =
  | { readonly success: true; readonly revision: number }
  | NativeEditFailure;

const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/**
 * Replace an atomic block's payload and regenerate its Word fallback in one native undo unit.
 * The block keeps its ID and control, so renderers and peers keep their identity. Atomic
 * blocks have no editable native text: this is their only content write. Suggestion mode
 * refuses, because a tracked atomic replacement belongs to the review executors.
 */
export function updateAtomicBlock(
  editor: DocxEditorInstance,
  update: AtomicBlockUpdate
): AtomicBlockUpdateResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (
    update.expectedRevision !== undefined &&
    update.expectedRevision !== surface.session.packageRevision()
  )
    return fail('stale-revision', 'Read the current document before changing this block.');
  if (surface.editingMode() === 'suggest')
    return fail('unsupported-operation', 'Tracked block changes require a review transaction.');
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'The document is read-only.');
  const pkg = surface.session.currentPackage();
  const block = buildBlockIndex(pkg).byId.get(update.blockId);
  if (!block) return fail('invalid-input', 'The block no longer exists.');
  const registry = activeMdxRegistry();
  const shown = registry.components.get(block.payload.type)?.wordFallback;
  if (registry.blockTags.get(block.payload.type)?.kind !== 'atomic' || !shown)
    return fail('unsupported-operation', 'Only atomic blocks with a Word fallback update here.');
  const payload: BlockPayload = { ...block.payload, attributes: { ...update.attributes } };
  let fallback: ReturnType<typeof shown>;
  try {
    fallback = shown(payload);
  } catch (error) {
    if (error instanceof MdxRefusal) return fail('invalid-input', error.message);
    throw error;
  }
  // An inline control holds runs, not paragraphs: its fallback is part of a sentence.
  const live = findNode(surface.session.part(), block.controlId);
  const liveContent =
    live?.kind === 'contentControl'
      ? live.children.find((child) => child.kind === 'contentControlContent')
      : undefined;
  if (
    !liveContent ||
    liveContent.kind === 'textValue' ||
    liveContent.children.some((child) => child.kind === 'run' || child.kind === 'hyperlink')
  )
    return fail('unsupported-operation', 'Only block-level atomic controls update here.');
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') return fail('invalid-document', 'The block payloads are missing.');
  const label = store.entries.get(block.payload.id)?.label ?? block.payload.type;
  const tag = blockTag(block.payload.id);
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${fallback.lines
      .map((line) => `<w:p>${runXml(line, fallback.code ? ['code'] : [])}</w:p>`)
      .join('')}</w:body></w:document>`,
    { name: '/sciobot-update-block.xml', contentType: surface.session.part().contentType }
  );
  const body = parsed.ok ? bodyStoryRoot(parsed.part) : null;
  if (!body || body.kind === 'textValue' || !body.children.length)
    return fail('invalid-input', 'The block fallback was refused.');
  const template = body.children;
  let committed = false;

  const replace = (working: OoxmlPackage): OoxmlPackage => {
    let next = working;
    if (fallback.code) {
      const styled = withInlineCodeStyle(next);
      if (!styled) throw new Error('The document cannot gain the inline code style.');
      next = styled;
    }
    const part = next.parts.get(surface.session.part().name);
    if (!part) throw new Error('The document part disappeared during the update.');
    const control = contentControlsIn(part.root).find(
      ({ node }) => contentControlPropertiesOf(node).tag === tag
    )?.node;
    const content = control?.children.find((child) => child.kind === 'contentControlContent');
    if (!content) throw new Error('The block control is missing.');
    const allocate = createNodeIdAllocator(part, 'new');
    const identity = {
      prefix: w14RootPrefix(part.root),
      used: new Set(usedParaIds(part.root)),
      seed: `scio-block-update:${block.payload.id}:${String(surface.session.packageRevision())}`,
    };
    const replaced = replaceChildren(
      part,
      content.id,
      template.map((node) => withFreshIds(node, allocate, identity))
    );
    if (!replaced.ok) throw new Error('The native store refused the block fallback.');
    const stored = withPayloadRecord(withPart(next, replaced.part), payload, label);
    if (!stored.ok) throw new Error(`Block payload storage refused: ${stored.reason}`);
    committed = true;
    return stored.pkg;
  };

  try {
    replace(pkg);
    committed = false;
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [replace], false, true);
    if (result.rejected || !result.committed || !committed)
      return fail('apply-failed', String(result.reason ?? 'The block update did not commit.'));
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The block update failed.'
    );
  }
  return { success: true, revision: surface.session.packageRevision() };
}
