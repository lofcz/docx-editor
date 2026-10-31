import { activeStorage } from '../registry/active';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  contentControlPropertiesOf,
  createNodeIdAllocator,
  escapeXml,
  findNode,
  insertChildren,
  readOoxmlPart,
  usedParaIds,
  w14RootPrefix,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { blockTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { withFreshIds } from './fresh-ids';
import type { BlockPayload } from './payload-schema';

export interface BlockInsertion {
  /** Registered container that receives the block, or null for the document body. */
  readonly parentControlId: string | null;
  /** Existing direct child of that container to insert before; null appends. */
  readonly beforeId: string | null;
  /** A new persistent block; its ID must not exist in the payload store or any control tag. */
  readonly payload: BlockPayload;
  /** Readable control alias and payload record label. */
  readonly label: string;
  /** Word-readable generated label, kept in the storage label control the projector removes. */
  readonly generatedLabel?: string;
  /** Initial editable paragraphs. Defaults to one empty paragraph, which receives the caret. */
  readonly paragraphs?: readonly string[];
  readonly expectedRevision?: number;
}

export type BlockInsertResult =
  | {
      readonly success: true;
      readonly revision: number;
      readonly controlId: string;
      readonly caret: { readonly paragraphId: string; readonly offset: number };
    }
  | NativeEditFailure;

const MAX_PARAGRAPHS = 1_000;
const MAX_TEXT = 1024 * 1024;

const failure = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

const run = (text: string) =>
  text ? `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>` : '';

function blockXml(insertion: BlockInsertion, paragraphs: readonly string[]): string {
  const label =
    insertion.generatedLabel === undefined
      ? ''
      : `<w:sdt><w:sdtPr><w:tag w:val="${activeStorage().labelTag}"/></w:sdtPr><w:sdtContent>` +
        `<w:p>${run(insertion.generatedLabel)}</w:p></w:sdtContent></w:sdt>`;
  return (
    `<w:sdt><w:sdtPr><w:alias w:val="${escapeXml(insertion.label)}"/>` +
    `<w:tag w:val="${escapeXml(blockTag(insertion.payload.id))}"/></w:sdtPr><w:sdtContent>` +
    label +
    paragraphs.map((text) => `<w:p>${run(text)}</w:p>`).join('') +
    '</w:sdtContent></w:sdt>'
  );
}

/** Paragraphs that are not inside the generated label, in document order. */
function editableParagraphs(control: OoxmlNode): string[] {
  const found: string[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (
      node.kind === 'contentControl' &&
      contentControlPropertiesOf(node).tag === activeStorage().labelTag
    )
      return;
    if (node.kind === 'paragraph') found.push(node.id);
    else node.children.forEach(visit);
  };
  if (control.kind !== 'textValue') control.children.forEach(visit);
  return found;
}

function findControl(root: OoxmlNode, tag: string): OoxmlNode | null {
  if (root.kind === 'textValue') return null;
  if (root.kind === 'contentControl' && contentControlPropertiesOf(root).tag === tag) return root;
  for (const child of root.children) {
    const found = findControl(child, tag);
    if (found) return found;
  }
  return null;
}

/** The node that holds a container's children: the body, or a control's `w:sdtContent`. */
function containerOf(part: OoxmlPart, parentControlId: string | null): OoxmlElement | string {
  if (parentControlId === null) {
    const body = bodyStoryRoot(part);
    return body && body.kind !== 'textValue' ? body : 'The document has no body.';
  }
  const control = findNode(part, parentControlId);
  if (!control || control.kind !== 'contentControl') return 'The parent block is missing.';
  const properties = contentControlPropertiesOf(control);
  if (properties.dataBinding) return 'The parent block is bound to custom XML.';
  if (properties.lock === 'contentLocked' || properties.lock === 'sdtContentLocked')
    return 'The parent block is locked.';
  const content = control.children.find((child) => child.kind === 'contentControlContent');
  if (!content) return 'The parent block has no content.';
  // A block-level container holds paragraphs; an inline control cannot take a block.
  if (content.children.some((child) => child.kind === 'run' || child.kind === 'hyperlink'))
    return 'The parent block is inline.';
  return content as OoxmlElement;
}

/**
 * Insert one registered container with its payload record, Word fallback, and caret paragraph
 * in a single native transaction. One undo removes all of them; redo restores them together.
 * Suggestion mode refuses: tracked block insertion belongs to the review executors.
 */
export function insertBlock(
  editor: DocxEditorInstance,
  insertion: BlockInsertion
): BlockInsertResult {
  const surface = editor.surface;
  if (!surface) return failure('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  const revision = surface.session.packageRevision();
  if (insertion.expectedRevision !== undefined && insertion.expectedRevision !== revision)
    return failure('stale-revision', 'Read the current document before inserting this block.');
  if (surface.editingMode() === 'suggest')
    return failure(
      'unsupported-operation',
      'Tracked block insertion requires a review transaction.'
    );
  const paragraphs = insertion.paragraphs ?? [''];
  if (
    !paragraphs.length ||
    paragraphs.length > MAX_PARAGRAPHS ||
    paragraphs.some(
      (text) => typeof text !== 'string' || text.length > MAX_TEXT || /[\r\n]/.test(text)
    ) ||
    (insertion.generatedLabel !== undefined && /[\r\n]/.test(insertion.generatedLabel))
  )
    return failure('invalid-input', 'A new block needs one to 1000 single-line paragraphs.');
  let tag: string;
  try {
    tag = blockTag(insertion.payload.id);
  } catch {
    return failure('invalid-input', 'The block ID is invalid.');
  }
  const pkg = surface.session.currentPackage();
  const part = surface.session.part();
  const store = readPayloadStore(pkg);
  if (store.status === 'invalid') return failure('invalid-document', store.reason);
  if (
    (store.status === 'ready' && store.entries.has(insertion.payload.id)) ||
    findControl(part.root, tag) !== null
  )
    return failure('invalid-input', 'A block with this ID already exists.');
  const container = containerOf(part, insertion.parentControlId);
  if (typeof container === 'string') return failure('invalid-input', container);
  if (
    insertion.beforeId !== null &&
    !container.children.some((child) => child.id === insertion.beforeId)
  )
    return failure('invalid-input', 'The insertion anchor is not a child of the parent block.');

  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${blockXml(insertion, paragraphs)}</w:body></w:document>`,
    { name: '/sciobot-insert-block.xml', contentType: part.contentType }
  );
  if (!parsed.ok) return failure('invalid-input', `The new block was refused: ${parsed.reason}.`);
  const body = bodyStoryRoot(parsed.part);
  if (!body || body.kind === 'textValue' || body.children.length !== 1)
    return failure('apply-failed', 'The new block did not produce one native control.');
  const template = body.children[0]!;
  let placedOnce = false;

  const place = (working: OoxmlPackage): OoxmlPackage => {
    const live = working.parts.get(part.name);
    if (!live) throw new Error('The document part disappeared during the insertion.');
    const target = containerOf(live, insertion.parentControlId);
    if (typeof target === 'string') throw new Error(target);
    const terminal = target.children.findIndex(
      (child) => child.kind !== 'textValue' && child.localName === 'sectPr'
    );
    const at =
      insertion.beforeId === null
        ? terminal < 0
          ? target.children.length
          : terminal
        : target.children.findIndex((child) => child.id === insertion.beforeId);
    if (at < 0) throw new Error('The insertion anchor disappeared during the insertion.');
    const node = withFreshIds(template, createNodeIdAllocator(live, 'new'), {
      prefix: w14RootPrefix(live.root),
      used: new Set(usedParaIds(live.root)),
      seed: `scio-block:${insertion.payload.id}`,
    });
    const placed = insertChildren(live, target.id, at, [node]);
    if (!placed.ok) throw new Error('The new block was refused by the native store.');
    if (!editableParagraphs(node).length)
      throw new Error('The new block has no editable paragraph.');
    const stored = withPayloadRecord(
      withPart(working, placed.part),
      insertion.payload,
      insertion.label
    );
    if (!stored.ok) throw new Error(`Block payload storage refused: ${stored.reason}`);
    placedOnce = true;
    return stored.pkg;
  };

  try {
    // Validate the complete candidate before the live transaction.
    place(pkg);
    placedOnce = false;
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [place], false, true);
    if (result.rejected || !result.committed || !placedOnce)
      return failure(
        'apply-failed',
        String(result.reason ?? 'The block insertion did not commit.')
      );
  } catch (error) {
    return failure(
      'apply-failed',
      error instanceof Error ? error.message : 'The block insertion failed.'
    );
  }
  // Collaboration can rename the new nodes during the commit, so read them back by tag.
  const live = findControl(surface.session.part().root, tag);
  const caret = live && editableParagraphs(live)[0];
  if (!live || !caret) return failure('apply-failed', 'The inserted block cannot be found.');
  const controlId = live.id;
  const position = { paragraphId: caret, offset: 0 };
  surface.setSelection({ anchor: position, head: position });
  return { success: true, revision: surface.session.packageRevision(), controlId, caret: position };
}
