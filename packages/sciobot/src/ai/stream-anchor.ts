import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  findNode,
  parentNodeOf,
  paraIdOf,
  storyParagraphs,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from '../blocks/block-index';

export interface StreamInsertionPoint {
  readonly parentId: string;
  /** Live sibling after which generation starts. */
  readonly afterId: string;
  /** Live following sibling, or null at the end of the container. */
  readonly beforeId: string | null;
}

export interface StreamAnchor {
  readonly state: 'active' | 'released';
  /** Resolve against current native content. A deleted scope releases the anchor. */
  resolve(): { readonly success: true; readonly point: StreamInsertionPoint } | NativeEditFailure;
  /** End the request without any document or history write. Idempotent. */
  release(): void;
}

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'stale-revision',
  error,
});

/**
 * A request-owned anchor after selected sibling blocks, or after the caret's body block.
 * No placeholder node enters the document. Registered scope identity follows native moves;
 * plain paragraphs use Word's stable paraId across collaboration rematerialization.
 */
export function createStreamAnchor(
  editor: DocxEditorInstance,
  selectedIds: readonly string[] = []
): { readonly success: true; readonly anchor: StreamAnchor } | NativeEditFailure {
  const surface = editor.surface;
  if (!surface)
    return {
      success: false,
      code: 'document-not-ready',
      error: 'Attach the document before starting generation.',
    };
  surface.flushPendingInput();
  const part = surface.session.part();
  const body = bodyStoryRoot(part);
  if (!body)
    return { success: false, code: 'invalid-document', error: 'The Word body is missing.' };
  let chosen: OoxmlNode | undefined;
  if (selectedIds.length) {
    if (new Set(selectedIds).size !== selectedIds.length)
      return { success: false, code: 'invalid-input', error: 'Selected blocks must be distinct.' };
    const nodes = selectedIds.map((id) => findNode(part, id));
    const parent = nodes[0] && parentNodeOf(part, nodes[0].id);
    if (!parent || nodes.some((node) => !node || parentNodeOf(part, node.id)?.id !== parent.id))
      return {
        success: false,
        code: 'invalid-input',
        error: 'Selected blocks must be live siblings.',
      };
    chosen = parent.children.filter((node) => selectedIds.includes(node.id)).at(-1);
  } else {
    chosen = findNode(part, surface.state().selection.head.paragraphId) ?? undefined;
    while (chosen && parentNodeOf(part, chosen.id)?.id !== body.id)
      chosen = parentNodeOf(part, chosen.id) ?? undefined;
  }
  if (!chosen || !['paragraph', 'contentControl', 'table'].includes(chosen.kind))
    return {
      success: false,
      code: 'invalid-input',
      error: 'Select a native block before starting generation.',
    };
  const chosenParent = parentNodeOf(part, chosen.id);
  if (!chosenParent || !['body', 'sdtContent', 'tc'].includes(chosenParent.localName))
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'This selection needs a native block insertion scope.',
    };
  const registeredId = [...buildBlockIndex(surface.session.currentPackage()).byId.values()].find(
    (block) => block.partName === part.name && block.controlId === chosen!.id
  )?.id;
  const paragraphIdentity = chosen.kind === 'paragraph' ? paraIdOf(chosen) : null;
  // Other structural roots retain their canonical node identity for this live part. A
  // replacement/load ends the anchor; do not guess identity from descendant text or position.
  const originalId = chosen.id;
  let active = true;
  let stop = () => {};
  function release(): void {
    if (!active) return;
    active = false;
    stop();
  }
  const anchor: StreamAnchor = {
    get state() {
      return active ? 'active' : 'released';
    },
    release,
    resolve() {
      if (!active || editor.surface !== surface) {
        release();
        return fail('The generation anchor has ended.');
      }
      const current = surface.session.part();
      if (current.name !== part.name) {
        release();
        return fail('The generation story changed.');
      }
      const block = registeredId
        ? buildBlockIndex(surface.session.currentPackage()).byId.get(registeredId)
        : undefined;
      const currentBody = bodyStoryRoot(current);
      const node = registeredId
        ? block?.partName === current.name
          ? findNode(current, block.controlId)
          : null
        : paragraphIdentity
          ? currentBody &&
            storyParagraphs(currentBody).find((node) => paraIdOf(node) === paragraphIdentity)
          : findNode(current, originalId);
      const parent = node && parentNodeOf(current, node.id);
      if (!node || !parent || !['body', 'sdtContent', 'tc'].includes(parent.localName)) {
        release();
        return fail('The generation scope was removed or changed.');
      }
      const index = parent.children.findIndex((child) => child.id === node.id);
      const next = parent.children[index + 1];
      return {
        success: true,
        point: {
          parentId: parent.id,
          afterId: node.id,
          beforeId:
            !next || (next.kind !== 'textValue' && next.localName === 'sectPr') ? null : next.id,
        },
      };
    },
  };
  stop = surface.observePaint({
    afterPaint() {
      if (active) anchor.resolve();
    },
    destroyed: release,
  });
  return { success: true, anchor };
}
