import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import type { TargetSourceInput } from '../diff/target-source';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import type { DirectPlanContext } from './apply-direct';

export type PreparedDirectSource =
  | { readonly success: true; readonly input: TargetSourceInput; readonly reused: boolean }
  | NativeEditFailure;

interface Landed {
  readonly token: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly projection: TargetSourceInput['projection'];
}

/**
 * The working projection a direct edit landed, reused as the next call's planning source while the
 * editor still holds exactly that canonical package. Any other commit (typing, undo, a review
 * decision, a remote revision) advances the package revision or replaces the package, so the next
 * call reads a fresh projection and plans against the live document.
 */
export class DirectTargetBaseline {
  private landed: Landed | undefined;

  /** Planning input for the next direct call: the remembered landing when still exact. */
  async prepare(
    editor: DocxEditorInstance,
    context: DirectPlanContext,
    signal?: AbortSignal
  ): Promise<PreparedDirectSource> {
    const token = readEditorRevision(editor, context.documentId, context.epoch),
      session = editor.surface?.session;
    if (!token || !session)
      return {
        success: false,
        code: 'document-not-ready',
        error: 'The native editor is not attached.',
      };
    const pkg = session.currentPackage(),
      landed = this.landed;
    if (landed && sameRevision(landed.token, token) && landed.pkg === pkg)
      return {
        success: true,
        reused: true,
        input: { revision: token, pkg, projection: landed.projection },
      };
    this.landed = undefined;
    const view = await readWorkingView(session, token, signal ? { signal } : undefined);
    if (!view.success) return view;
    return {
      success: true,
      reused: false,
      input: { revision: token, pkg, projection: view.view.projection },
    };
  }

  /** Read and remember the landed working projection right after a committed direct edit. */
  async remember(
    editor: DocxEditorInstance,
    context: DirectPlanContext,
    signal?: AbortSignal
  ): Promise<PreparedDirectSource> {
    this.landed = undefined;
    const prepared = await this.prepare(editor, context, signal);
    if (!prepared.success) return prepared;
    const session = editor.surface!.session;
    // A commit during the asynchronous read leaves nothing exact to remember.
    if (session.currentPackage() === prepared.input.pkg)
      this.landed = {
        token: prepared.input.revision,
        pkg: prepared.input.pkg,
        projection: prepared.input.projection,
      };
    return prepared;
  }

  invalidate(): void {
    this.landed = undefined;
  }
}
