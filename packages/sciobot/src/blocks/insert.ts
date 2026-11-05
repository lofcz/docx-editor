import { activeStorage } from '../registry/active';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { SemanticSelection } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  bodyStoryRoot,
  contentControlPropertiesOf,
  createNodeIdAllocator,
  ensureHyperlinkRelationship,
  escapeXml,
  findNode,
  insertChildren,
  parentNodeOf,
  readOoxmlPart,
  usedParaIds,
  w14RootPrefix,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
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
  /**
   * An external URL the first paragraph links to, so Word readers can open a media block's
   * source. It must be an absolute URL that `sanitizeHref` admits.
   */
  readonly link?: string;
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

const R_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function blockXml(
  insertion: BlockInsertion,
  paragraphs: readonly string[],
  linkId: string | null
): string {
  const label =
    insertion.generatedLabel === undefined
      ? ''
      : `<w:sdt><w:sdtPr><w:tag w:val="${activeStorage().labelTag}"/></w:sdtPr><w:sdtContent>` +
        `<w:p>${run(insertion.generatedLabel)}</w:p></w:sdtContent></w:sdt>`;
  return (
    `<w:sdt><w:sdtPr><w:alias w:val="${escapeXml(insertion.label)}"/>` +
    `<w:tag w:val="${escapeXml(blockTag(insertion.payload.id))}"/></w:sdtPr><w:sdtContent>` +
    label +
    paragraphs
      .map((text, index) =>
        index === 0 && linkId !== null
          ? // The hyperlink declares its own `r` prefix: the target part may not.
            `<w:p><w:hyperlink xmlns:r="${R_NAMESPACE}" r:id="${escapeXml(linkId)}">${run(text)}</w:hyperlink></w:p>`
          : `<w:p>${run(text)}</w:p>`
      )
      .join('') +
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
  return commitBlockInsertion(editor, insertion);
}

/** Insert an atomic block at a captured text range, preserving both unselected paragraph sides. */
export function insertBlockAtSelection(
  editor: DocxEditorInstance,
  insertion: Omit<BlockInsertion, 'parentControlId' | 'beforeId'> & {
    readonly selection: SemanticSelection;
    readonly expectedRevision: number;
  }
): BlockInsertResult {
  return commitBlockInsertion(
    editor,
    { ...insertion, parentControlId: null, beforeId: null },
    insertion.selection
  );
}

function commitBlockInsertion(
  editor: DocxEditorInstance,
  insertion: BlockInsertion,
  replacement?: SemanticSelection
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
  const source = surface.session.part();
  let part = source;
  const ops: TreeDocOp[] = [];
  let landing: string | undefined;
  if (replacement) {
    if (surface.editingMode() !== 'edit' || editor.getTableCellSelection())
      return failure(
        'unsupported-operation',
        'Block insertion requires an editable text selection.'
      );
    const plan = surface.planRangeDeletion(replacement);
    const expanded =
      replacement.anchor.paragraphId !== replacement.head.paragraphId ||
      replacement.anchor.offset !== replacement.head.offset;
    if (expanded && !plan.ops.length)
      return failure('unsupported-operation', 'This text selection cannot be replaced.');
    ops.push(...plan.ops, {
      op: 'splitParagraph',
      paragraphId: plan.replaceAt.paragraphId,
      offset: plan.replaceAt.offset,
    });
    landing = plan.replaceAt.paragraphId;
    for (const op of ops) {
      const applied = applyTreeOp(part, op);
      if (!applied.ok)
        return failure('unsupported-operation', `Block insertion refused: ${applied.reason}.`);
      part = applied.part;
    }
  }
  const placement = (
    live: OoxmlPart
  ): { parentControlId: string | null; beforeId: string | null } => {
    if (!landing) return insertion;
    const parent = parentNodeOf(live, landing);
    const at = parent?.children.findIndex((child) => child.id === landing) ?? -1;
    const tail = parent?.children[at + 1];
    if (!parent || at < 0 || tail?.kind !== 'paragraph')
      throw Error('The insertion paragraph is missing.');
    if (parent.kind === 'body') return { parentControlId: null, beforeId: tail.id };
    const owner = parent.kind === 'contentControlContent' && parentNodeOf(live, parent.id);
    if (owner && owner.kind === 'contentControl')
      return { parentControlId: owner.id, beforeId: tail.id };
    throw Error('This container does not support atomic block insertion.');
  };
  const store = readPayloadStore(pkg);
  if (store.status === 'invalid') return failure('invalid-document', store.reason);
  if (
    (store.status === 'ready' && store.entries.has(insertion.payload.id)) ||
    findControl(part.root, tag) !== null
  )
    return failure('invalid-input', 'A block with this ID already exists.');
  let destination: ReturnType<typeof placement>;
  try {
    destination = placement(part);
  } catch (error) {
    return failure('unsupported-operation', String(error));
  }
  const container = containerOf(part, destination.parentControlId);
  if (typeof container === 'string') return failure('invalid-input', container);
  if (
    destination.beforeId !== null &&
    !container.children.some((child) => child.id === destination.beforeId)
  )
    return failure('invalid-input', 'The insertion anchor is not a child of the parent block.');

  if (insertion.link !== undefined && !paragraphs[0])
    return failure('invalid-input', 'A linked block needs text in its first paragraph.');
  const templateFor = (linkId: string | null): OoxmlNode | string => {
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${blockXml(insertion, paragraphs, linkId)}</w:body></w:document>`,
      { name: '/sciobot-insert-block.xml', contentType: part.contentType }
    );
    if (!parsed.ok) return `The new block was refused: ${parsed.reason}.`;
    const body = bodyStoryRoot(parsed.part);
    if (!body || body.kind === 'textValue' || body.children.length !== 1)
      return 'The new block did not produce one native control.';
    return body.children[0]!;
  };
  const unlinked = templateFor(null);
  if (typeof unlinked === 'string') return failure('invalid-input', unlinked);
  if (insertion.link !== undefined && !ensureHyperlinkRelationship(pkg, insertion.link, part.name))
    return failure('invalid-input', 'The block link is not an allowed absolute URL.');
  let placedOnce = false;

  const place = (start: OoxmlPackage): OoxmlPackage => {
    let working = start;
    let template = unlinked;
    if (insertion.link !== undefined) {
      // The relationship ID depends on the package the transform runs on.
      const ensured = ensureHyperlinkRelationship(working, insertion.link, part.name);
      if (!ensured) throw new Error('The block link relationship was refused.');
      working = ensured.pkg;
      const linked = templateFor(ensured.relationshipId);
      if (typeof linked === 'string') throw new Error(linked);
      template = linked;
    }
    const live = working.parts.get(part.name);
    if (!live) throw new Error('The document part disappeared during the insertion.');
    const point = placement(live);
    const target = containerOf(live, point.parentControlId);
    if (typeof target === 'string') throw new Error(target);
    const terminal = target.children.findIndex(
      (child) => child.kind !== 'textValue' && child.localName === 'sectPr'
    );
    const at =
      point.beforeId === null
        ? terminal < 0
          ? target.children.length
          : terminal
        : target.children.findIndex((child) => child.id === point.beforeId);
    if (at < 0) throw new Error('The insertion anchor disappeared during the insertion.');
    const node = withFreshIds(template, createNodeIdAllocator(live, 'new'), {
      prefix: w14RootPrefix(live.root),
      used: new Set(usedParaIds(live.root)),
      seed: `scio-block:${insertion.payload.id}`,
    });
    const placed = insertChildren(live, target.id, at, [node]);
    if (!placed.ok)
      throw new Error(
        `The new block was refused by the native store (${placed.issues.map((issue) => issue.code).join(', ')}).`
      );
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
    place(replacement ? withPart(pkg, part) : pkg);
    placedOnce = false;
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [place],
      false,
      true,
      replacement ? 'after' : 'before'
    );
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
  const liveParent = live && parentNodeOf(surface.session.part(), live.id);
  const tail =
    liveParent &&
    liveParent.children[liveParent.children.findIndex((child) => child.id === live!.id) + 1];
  const caret = replacement
    ? tail?.kind === 'paragraph'
      ? tail.id
      : null
    : live && editableParagraphs(live)[0];
  if (!live || !caret) return failure('apply-failed', 'The inserted block cannot be found.');
  const controlId = live.id;
  const position = { paragraphId: caret, offset: 0 };
  surface.setSelection({ anchor: position, head: position });
  return { success: true, revision: surface.session.packageRevision(), controlId, caret: position };
}
