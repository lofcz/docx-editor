import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import type { DirectChangeLedger, DirectChangeStatus } from './change-ledger';

/** One paintable highlight: a whole native paragraph that a recorded direct change wrote. */
export interface DirectHighlight {
  readonly changeId: string;
  readonly status: DirectChangeStatus;
  readonly paragraphId: string;
  readonly start: number;
  readonly end: number;
}

/**
 * Resolve ledger changes to live paragraph ranges for painting. Pending playback changes stay
 * hidden. Every range addresses a paragraph that exists in the current canonical part; a change
 * whose nodes vanished (undo, deletion, a remote edit) has already been dropped by the ledger,
 * and any remaining stale node is skipped rather than painted.
 */
export function directHighlights(
  editor: DocxEditorInstance,
  ledger: DirectChangeLedger
): readonly DirectHighlight[] {
  const surface = editor.surface;
  if (!surface) return [];
  // Queued typing commits first, so the ledger settles touched changes before painting.
  surface.flushPendingInput();
  const part = surface.session.part(),
    out: DirectHighlight[] = [];
  for (const change of ledger.changes) {
    if (change.pending) continue;
    for (const paragraphId of change.nodeIds) {
      if (findNode(part, paragraphId)?.kind !== 'paragraph') continue;
      const text = paragraphTextOf(part, paragraphId) ?? '';
      out.push(
        Object.freeze({
          changeId: change.id,
          status: change.status,
          paragraphId,
          start: 0,
          end: text.length,
        })
      );
    }
  }
  return Object.freeze(out);
}
