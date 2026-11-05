import type { EditorCollaborationSession } from '../collaboration/index.ts';
import { findNode, type OoxmlPackage } from '../store/index.ts';
import type { DrawingSelectionIntent } from './paginated-surface-state.ts';

/** Adopt an exact replica identity while its old local alias still exists. */
export function remapDrawingIntent(
  intent: DrawingSelectionIntent,
  collaboration: EditorCollaborationSession | undefined,
  pkg: OoxmlPackage
): DrawingSelectionIntent {
  if (intent.kind !== 'pointer' || !collaboration?.resolveNodeId) return intent;
  const drawingNodeId = collaboration.resolveNodeId(intent.drawingNodeId);
  if (drawingNodeId === intent.drawingNodeId) return intent;
  for (const part of pkg.parts.values()) {
    // Eager atomic publication can retain the ordinary drawing's local address.
    if (findNode(part, intent.drawingNodeId)) return intent;
  }
  for (const part of pkg.parts.values())
    if (findNode(part, drawingNodeId)) return { kind: 'pointer', drawingNodeId };
  return intent;
}
