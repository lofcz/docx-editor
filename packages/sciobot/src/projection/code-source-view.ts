import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import { findNode, type OoxmlPackage } from '@docx-editor.dev/core/store';
import type { IndexedBlock } from '../blocks/block-index';
import { nativeCodeSource } from '../blocks/code-native-source';

/** Read the requested revision half from native code, without writing its payload cache. */
export function codeBlockForView(
  pkg: OoxmlPackage,
  block: IndexedBlock,
  displayMode: ExportSemanticLayout['reviewDisplayMode']
): IndexedBlock | null {
  if (block.payload.type !== 'code_block' || block.payload.version !== 2) return block;
  const part = pkg.parts.get(block.partName);
  const control = part && findNode(part, block.controlId);
  const view =
    displayMode === 'proposed' || displayMode === 'simple-markup' ? 'working' : 'committed';
  const source = part && control && nativeCodeSource(part, control, view);
  if (!source) return null;
  return {
    ...block,
    payload: {
      ...block.payload,
      attributes: { ...block.payload.attributes, value: source.value },
    },
  };
}
