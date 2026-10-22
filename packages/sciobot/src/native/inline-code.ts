import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { segmentGraphemes } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  createNodeIdAllocator,
  DEFAULT_FORMATTING_DISPLAY_MODE,
  findNode,
  insertChildren,
  paragraphTextOf,
  partNameKey,
  propertyContainer,
  removeNode,
  readOoxmlPart,
  runPropertyEdits,
  runsCovering,
  stylesPartOf,
  withNewPart,
  withPart,
  withRelationship,
  type FormattingDisplayMode,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type OoxmlProperty,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { readEditorRevision } from '../projection/revision-cache';

/**
 * The character style that identifies an application-authored code span.
 *
 * The identity is the run's own `w:rStyle`, never its font: imported monospaced text stays
 * ordinary text. Word keeps a used custom character style through save, so the identity survives
 * a round trip through desktop Word.
 */
export const SCIOBOT_CODE_STYLE_ID = 'SciobotCode';
/** Readable name shown in Word's style gallery. */
export const SCIOBOT_CODE_STYLE_NAME = 'Sciobot Code';
/** Style IDs the private Markdown projection reads as code spans. */
export const SCIOBOT_CODE_CHARACTER_STYLES: readonly string[] = Object.freeze([
  SCIOBOT_CODE_STYLE_ID,
]);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STYLES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const DERIVED_PREFIX = `${SCIOBOT_CODE_STYLE_ID}__`;

/** Reversible identity: a derived code style retains the exact imported character style. */
export function inlineCodeStyleId(base: string | undefined): string {
  return base === undefined ? SCIOBOT_CODE_STYLE_ID : DERIVED_PREFIX + encodeURIComponent(base);
}

function codeStyleBase(id: string): string | undefined {
  if (!id.startsWith(DERIVED_PREFIX)) return undefined;
  try {
    const base = decodeURIComponent(id.slice(DERIVED_PREFIX.length));
    return base && inlineCodeStyleId(base) === id ? base : undefined;
  } catch {
    return undefined;
  }
}

const xmlAttribute = (value: string): string =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

function styleDefinition(id: string): string {
  const base = codeStyleBase(id);
  return (
    `<w:style w:type="character" w:customStyle="1" w:styleId="${xmlAttribute(id)}">` +
    `<w:name w:val="${SCIOBOT_CODE_STYLE_NAME}"/>` +
    (base === undefined ? '' : `<w:basedOn w:val="${xmlAttribute(base)}"/>`) +
    '<w:uiPriority w:val="99"/>' +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas" w:cs="Consolas"/>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:rPr></w:style>'
  );
}

export interface InlineCodeRequest {
  readonly paragraphId: string;
  /** UTF-16 native paragraph offsets on grapheme boundaries. */
  readonly start: number;
  readonly end: number;
  /** `true` marks the range as code; `false` removes the code identity. */
  readonly code: boolean;
}

export type InlineCodeOpsResult =
  | {
      readonly success: true;
      readonly ops: readonly TreeDocOp[];
      /** The package needs its code style definition added or deduplicated with these operations. */
      readonly needsStyle: boolean;
    }
  | NativeEditFailure;

export interface InlineCodeContext {
  readonly documentId: string;
  readonly epoch: string;
  /** The revision the caller read; a different live revision refuses. */
  readonly revision: RevisionToken;
}

export type InlineCodeResult =
  | { readonly success: true; readonly revision: RevisionToken; readonly committed: boolean }
  | NativeEditFailure;

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

function attribute(node: OoxmlNode, localName: string): string | undefined {
  if (node.kind === 'textValue') return undefined;
  return node.attributes.find((entry) => entry.namespaceUri === W && entry.localName === localName)
    ?.value;
}

function isWml(node: OoxmlNode, localName: string): node is OoxmlElement {
  return node.kind !== 'textValue' && node.namespaceUri === W && node.localName === localName;
}

/** The run's own `w:rStyle` reference, or undefined. */
export function runCharacterStyle(run: OoxmlNode): string | undefined {
  const container = propertyContainer(run, 'runProperties', 'rPr');
  if (!container || container.kind === 'textValue') return undefined;
  const reference = container.children.find((child) => isWml(child, 'rStyle'));
  return reference ? attribute(reference, 'val') : undefined;
}

/** Whether the run carries the explicit application code identity. */
export function isInlineCodeRun(run: OoxmlNode, pkg?: OoxmlPackage): boolean {
  const style = runCharacterStyle(run);
  return (
    style !== undefined &&
    (pkg ? inlineCodeCharacterStyles(pkg) : SCIOBOT_CODE_CHARACTER_STYLES).includes(style)
  );
}

function styleDefinitions(pkg: OoxmlPackage): readonly OoxmlElement[] {
  const root = stylesPartOf(pkg)?.root;
  const styles: OoxmlElement[] = [];
  for (const child of root?.children ?? []) if (isWml(child, 'style')) styles.push(child);
  return styles;
}

function codeStyleDefinitions(
  pkg: OoxmlPackage,
  id = SCIOBOT_CODE_STYLE_ID
): readonly OoxmlElement[] {
  return styleDefinitions(pkg).filter((style) => attribute(style, 'styleId') === id);
}

/** Only explicit code definitions qualify; a similarly named imported style stays ordinary. */
export function inlineCodeCharacterStyles(pkg: OoxmlPackage): readonly string[] {
  const candidates = new Map<string, boolean>();
  for (const style of styleDefinitions(pkg)) {
    const id = attribute(style, 'styleId');
    if (id === undefined || codeStyleBase(id) === undefined) continue;
    candidates.set(id, candidates.get(id) !== false && matchesCodeDefinition(style, id));
  }
  return [SCIOBOT_CODE_STYLE_ID, ...[...candidates].flatMap(([id, valid]) => (valid ? [id] : []))];
}

function matchesCodeDefinition(style: OoxmlElement, id: string): boolean {
  const base = codeStyleBase(id);
  return (
    attribute(style, 'type') === 'character' &&
    (base === undefined ||
      (style.children.some(
        (child) => isWml(child, 'basedOn') && attribute(child, 'val') === base
      ) &&
        style.children.some(
          (child) => isWml(child, 'name') && attribute(child, 'val') === SCIOBOT_CODE_STYLE_NAME
        )))
  );
}

/**
 * `present`, `missing`, `duplicate` when concurrent peers each added the definition, or
 * `conflict` when the ID names a style of another type.
 */
export function inlineCodeStyleState(
  pkg: OoxmlPackage,
  id = SCIOBOT_CODE_STYLE_ID
): 'present' | 'missing' | 'duplicate' | 'conflict' {
  const base = codeStyleBase(id);
  if (id !== SCIOBOT_CODE_STYLE_ID && base === undefined) return 'conflict';
  const matches = codeStyleDefinitions(pkg, id);
  if (matches.length === 0) return 'missing';
  if (!matches.every((style) => matchesCodeDefinition(style, id))) return 'conflict';
  return matches.length === 1 ? 'present' : 'duplicate';
}

/**
 * Seed definitions for existing character styles before room bootstrap. Creating or deduplicating
 * these definitions in an actor's undoable edit can remove a definition another actor still uses.
 * Conflicting derived identities stay untouched; the code command reports that conflict.
 */
export function inlineCodeInitializationStyles(pkg: OoxmlPackage): readonly string[] {
  const codeStyles = new Set(inlineCodeCharacterStyles(pkg));
  const ids = new Set([SCIOBOT_CODE_STYLE_ID]);
  for (const style of styleDefinitions(pkg)) {
    const base = attribute(style, 'styleId');
    if (attribute(style, 'type') !== 'character' || base === undefined || codeStyles.has(base))
      continue;
    const id = inlineCodeStyleId(base);
    if (id.length <= 512 && inlineCodeStyleState(pkg, id) !== 'conflict') ids.add(id);
  }
  return [...ids];
}

/** The document's own Hyperlink character style, which link runs return to without code. */
export function hyperlinkCharacterStyle(pkg: OoxmlPackage): string | null {
  for (const style of styleDefinitions(pkg)) {
    if (attribute(style, 'type') !== 'character') continue;
    const name = style.children.find((child) => isWml(child, 'name'));
    if (name && attribute(name, 'val')?.trim().toLowerCase() === 'hyperlink')
      return attribute(style, 'styleId') ?? null;
  }
  return null;
}

function withFreshIds(node: OoxmlNode, nextId: () => string): OoxmlNode {
  if (node.kind === 'textValue') return Object.freeze({ ...node, id: nextId() });
  return Object.freeze({
    ...node,
    id: nextId(),
    children: node.children.map((child) => withFreshIds(child, nextId)),
  }) as OoxmlElement;
}

function parsedStyles(name: string, inner: string): OoxmlPart | null {
  const read = readOoxmlPart(`<w:styles xmlns:w="${W}">${inner}</w:styles>`, {
    name,
    contentType: STYLES_CONTENT_TYPE,
  });
  return read.ok ? read.part : null;
}

/**
 * Add the code character style definition when the package lacks it, or keep only the first of
 * several definitions that offline peers added concurrently.
 *
 * Returns the package unchanged when exactly one definition exists. Returns null when the ID names
 * a style of another type or the package cannot gain a styles part; callers refuse before writing.
 */
export function withInlineCodeStyle(
  pkg: OoxmlPackage,
  ids: readonly string[] = SCIOBOT_CODE_CHARACTER_STYLES
): OoxmlPackage | null {
  let working = pkg;
  for (const id of new Set(ids)) {
    const next = withOneCodeStyle(working, id);
    if (!next) return null;
    working = next;
  }
  return working;
}

function withOneCodeStyle(pkg: OoxmlPackage, id: string): OoxmlPackage | null {
  const state = inlineCodeStyleState(pkg, id);
  if (state === 'present') return pkg;
  if (state === 'conflict') return null;
  const part = stylesPartOf(pkg);
  if (state === 'duplicate') {
    let next = part!;
    for (const extra of codeStyleDefinitions(pkg, id).slice(1)) {
      const removed = removeNode(next, extra.id);
      if (!removed.ok) return null;
      next = removed.part;
    }
    return withPart(pkg, next);
  }
  if (part) {
    const template = parsedStyles(part.name, styleDefinition(id));
    const definition = template?.root.children[0];
    if (!definition) return null;
    const inserted = insertChildren(part, part.root.id, part.root.children.length, [
      withFreshIds(definition, createNodeIdAllocator(part)),
    ]);
    return inserted.ok ? withPart(pkg, inserted.part) : null;
  }
  const main = pkg.mainDocumentPart;
  const directory = main.slice(0, main.lastIndexOf('/') + 1);
  const occupied = new Set([...pkg.parts.keys(), ...pkg.partBytes.keys()].map(partNameKey));
  let filename = 'styles.xml';
  for (let suffix = 1; occupied.has(partNameKey(directory + filename)); suffix += 1)
    filename = `sciobot-styles-${suffix}.xml`;
  const name = directory + filename;
  const created = parsedStyles(name, styleDefinition(id));
  if (!created || pkg.parts.has(name)) return null;
  const withStyles = withNewPart(pkg, name, created.root as OoxmlElement, STYLES_CONTENT_TYPE);
  if (!withStyles.parts.has(name)) return null;
  const related = withRelationship(withStyles, main, STYLES_RELATIONSHIP, filename);
  return related.ok ? related.pkg : null;
}

/** IDs of runs whose nearest inline wrapper chain includes a `w:hyperlink`. */
function linkedRunIds(paragraph: OoxmlNode): ReadonlySet<string> {
  const ids = new Set<string>();
  const walk = (node: OoxmlNode, linked: boolean): void => {
    if (node.kind === 'textValue') return;
    const inside = linked || isWml(node, 'hyperlink');
    if (node.kind === 'run') {
      if (inside) ids.add(node.id);
      return;
    }
    for (const child of node.children) walk(child, inside);
  };
  walk(paragraph, false);
  return ids;
}

/** One native run slice whose character style changes for a code request. */
export interface InlineCodeRunChange {
  readonly start: number;
  readonly end: number;
  /** The run's current `w:rStyle`, or undefined. */
  readonly before: string | undefined;
  /** The new `w:rStyle`, or null to remove it. */
  readonly after: string | null;
  /** The run's own authored property bag, unchanged by the request. */
  readonly properties: readonly OoxmlProperty[];
  readonly targetRunIds?: readonly string[];
}

export type InlineCodeRunChangesResult =
  | { readonly success: true; readonly changes: readonly InlineCodeRunChange[] }
  | NativeEditFailure;

/**
 * Resolve the per-run character style changes for one paragraph range, without writing.
 *
 * Each run keeps its own authored property bag; only `w:rStyle` changes. A link run gives its
 * Hyperlink style to code and gets it back without code. Other imported character styles
 * use a derived code style with `w:basedOn`, so their definitions and effective marks survive.
 */
export function inlineCodeRunChanges(
  part: OoxmlPart,
  request: InlineCodeRequest,
  options: {
    readonly displayMode?: FormattingDisplayMode;
    /** The document's Hyperlink character style ID; see {@link hyperlinkCharacterStyle}. */
    readonly hyperlinkStyle?: string | null;
    readonly codeStyles?: readonly string[];
  } = {}
): InlineCodeRunChangesResult {
  if (
    !request ||
    typeof request.paragraphId !== 'string' ||
    typeof request.code !== 'boolean' ||
    !Number.isInteger(request.start) ||
    !Number.isInteger(request.end)
  )
    return fail('Invalid inline code request.');
  const displayMode = options.displayMode ?? DEFAULT_FORMATTING_DISPLAY_MODE;
  const paragraph = findNode(part, request.paragraphId);
  const text = paragraphTextOf(part, request.paragraphId);
  if (paragraph?.kind !== 'paragraph' || text === null)
    return fail('The inline code paragraph is not a native body paragraph.', 'ambiguous-target');
  if (request.start < 0 || request.end > text.length || request.start >= request.end)
    return fail('Invalid inline code range.');
  const boundaries = new Set([0, text.length]);
  for (const item of segmentGraphemes(text)) {
    boundaries.add(item.utf16From);
    boundaries.add(item.utf16To);
  }
  if (!boundaries.has(request.start) || !boundaries.has(request.end))
    return fail('An inline code range splits a native grapheme.', 'ambiguous-target');
  const linked = linkedRunIds(paragraph);
  const codeStyles = options.codeStyles ?? SCIOBOT_CODE_CHARACTER_STYLES;
  const changes: InlineCodeRunChange[] = [];
  for (const edit of runPropertyEdits(
    part,
    request.paragraphId,
    request.start,
    request.end,
    [],
    displayMode
  )) {
    const run = runsCovering(part, request.paragraphId, edit.start, edit.end, displayMode)[0];
    if (!run) continue;
    const before = runCharacterStyle(run);
    let after: string | null;
    if (request.code) {
      if (before !== undefined && codeStyles.includes(before)) continue;
      after = inlineCodeStyleId(before);
      if (after.length > 512)
        return fail(
          'The derived inline code style exceeds the native style ID limit.',
          'unsupported-operation'
        );
    } else {
      if (before === undefined || !codeStyles.includes(before)) continue;
      after =
        codeStyleBase(before) ?? (linked.has(run.id) ? (options.hyperlinkStyle ?? null) : null);
    }
    changes.push({
      start: edit.start,
      end: edit.end,
      before,
      after,
      properties: edit.properties,
      ...(edit.targetRunIds ? { targetRunIds: edit.targetRunIds } : {}),
    });
  }
  return { success: true, changes };
}

/**
 * Canonical operations that set or remove the code identity on one paragraph range.
 * See {@link inlineCodeRunChanges} for the per-run rules.
 */
export function inlineCodeOps(
  pkg: OoxmlPackage,
  request: InlineCodeRequest,
  displayMode: FormattingDisplayMode = DEFAULT_FORMATTING_DISPLAY_MODE
): InlineCodeOpsResult {
  const part = pkg.parts.get(pkg.mainDocumentPart);
  if (!part) return fail('The package has no main document part.', 'invalid-document');
  if (request?.code === true && inlineCodeStyleState(pkg) === 'conflict')
    return fail(
      `The document defines ${SCIOBOT_CODE_STYLE_ID} as a non-character style.`,
      'unsupported-operation'
    );
  const resolved = inlineCodeRunChanges(part, request, {
    displayMode,
    hyperlinkStyle: hyperlinkCharacterStyle(pkg),
    codeStyles: inlineCodeCharacterStyles(pkg),
  });
  if (!resolved.success) return resolved;
  const styles = resolved.changes.flatMap((change) =>
    change.after !== null && request.code ? [change.after] : []
  );
  if (styles.some((id) => inlineCodeStyleState(pkg, id) === 'conflict'))
    return fail('The document defines a conflicting inline code style.', 'unsupported-operation');
  const ops: TreeDocOp[] = resolved.changes.map((change) => ({
    op: 'setRunProperties',
    paragraphId: request.paragraphId,
    start: change.start,
    end: change.end,
    properties: change.properties,
    characterStyle: change.after,
    ...(change.targetRunIds ? { targetRunIds: change.targetRunIds } : {}),
  }));
  // Preview the complete batch so a refused operation never reaches the live document.
  let preview = part;
  for (const op of ops) {
    const result = applyTreeOp(preview, op);
    if (!result.ok) return fail(`Native ${op.op} refused: ${result.reason}.`, 'apply-failed');
    preview = result.part;
  }
  return {
    success: true,
    ops,
    needsStyle:
      ops.length > 0 &&
      (inlineCodeCharacterStyles(pkg).some((id) => inlineCodeStyleState(pkg, id) === 'duplicate') ||
        styles.some((id) => inlineCodeStyleState(pkg, id) === 'missing')),
  };
}

/**
 * Set or remove inline code through one gated native transaction with one undo step.
 *
 * Editing mode writes directly. Suggestion mode keeps tracking on, so the surface records each
 * run change as a Word formatting revision; the style definition itself is untracked package
 * data, as in Word.
 */
export function applyInlineCode(
  editor: DocxEditorInstance,
  request: InlineCodeRequest,
  context: InlineCodeContext
): InlineCodeResult {
  if (!context || typeof context.documentId !== 'string' || typeof context.epoch !== 'string')
    return fail('Invalid inline code context.');
  const surface = editor.surface;
  const actual = readEditorRevision(editor, context.documentId, context.epoch);
  if (!surface || !actual) return fail('The native editor is not attached.', 'document-not-ready');
  if (!context.revision || !sameRevision(actual, context.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before changing inline code.',
      ...(context.revision ? { expectedRevision: context.revision } : {}),
      actualRevision: actual,
    };
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  const pkg = surface.session.currentPackage();
  // The formatting projection of the review view: Simple and No Markup lay out the proposal.
  const view = surface.revisionDisplayMode();
  const displayMode: FormattingDisplayMode =
    view === 'original' ? 'original' : view === 'all-markup' ? 'all-markup' : 'proposed';
  const planned = inlineCodeOps(pkg, request, displayMode);
  if (!planned.success) return planned;
  if (!planned.ops.length) return { success: true, revision: actual, committed: false };
  const styles = [
    ...inlineCodeCharacterStyles(pkg),
    ...planned.ops.flatMap((op) =>
      op.op === 'setRunProperties' && typeof op.characterStyle === 'string' && request.code
        ? [op.characterStyle]
        : []
    ),
  ];
  if (planned.needsStyle && !withInlineCodeStyle(pkg, styles))
    return fail('The document cannot gain the inline code style.', 'unsupported-operation');
  const ensureStyle = (working: OoxmlPackage): OoxmlPackage => {
    const next = withInlineCodeStyle(working, styles);
    if (!next) throw new Error('The document cannot gain the inline code style.');
    return next;
  };
  try {
    const result = surface.applyAutomationOps(
      () => planned.ops,
      { kind: 'body' },
      planned.needsStyle ? [ensureStyle] : undefined,
      false,
      true
    );
    if (result.rejected || !result.committed)
      return {
        success: false,
        code: 'apply-failed',
        error: String(result.reason ?? 'The inline code change did not commit.'),
        expectedRevision: actual,
        actualRevision: { ...actual, revision: surface.session.packageRevision() },
      };
  } catch (error) {
    return {
      success: false,
      code: 'apply-failed',
      error: error instanceof Error ? error.message : String(error),
      expectedRevision: actual,
      actualRevision: { ...actual, revision: surface.session.packageRevision() },
    };
  }
  return {
    success: true,
    revision: { ...actual, revision: surface.session.packageRevision() },
    committed: true,
  };
}
