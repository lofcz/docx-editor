import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import type { TreeDocxSessionView } from '../binding/tree-session.ts';
import { ensureStyleDefinition } from '../store/package/style-definition-edit.ts';

/** Materialize these named styles on first use, including in a document without styles.xml. */
const presets = {
  BandedGridBlue: { name: 'Banded Grid Blue', dark: '294F79', medium: 'C1D5EC', pale: 'E5EFF8' },
  BandedGridGreen: { name: 'Banded Grid Green', dark: '376457', medium: 'C0DCCB', pale: 'E6F2EB' },
};
export function tableStyleDefinition(styleId: string | undefined) {
  if (styleId !== 'BandedGridBlue' && styleId !== 'BandedGridGreen') return undefined;
  const preset = presets[styleId];
  return ensureStyleDefinition(styleId, (el) => {
    const region = (type: string, fill: string, bold = false, white = false) =>
      el('tblStylePr', { type }, [
        ...(bold
          ? [el('rPr', {}, [el('b'), ...(white ? [el('color', { val: 'FFFFFF' })] : [])])]
          : []),
        el('tcPr', {}, [el('shd', { val: 'clear', fill })]),
      ]);
    return el('style', { type: 'table', styleId, customStyle: '1' }, [
      el('name', { val: preset.name }),
      region('band1Horz', preset.pale),
      region('band1Vert', preset.medium),
      region('firstCol', preset.medium, true),
      region('lastCol', preset.medium, true),
      region('firstRow', preset.dark, true, true),
      region('lastRow', preset.dark, true, true),
    ]);
  });
}

export function tableStyleEdits(ops: readonly TreeDocOp[], session: TreeDocxSessionView) {
  const op = ops.find((op) => op.op === 'setTableProperties' && op.styleId);
  if (
    op?.op !== 'setTableProperties' ||
    session.documentStyles().some((s) => s.styleId === op.styleId)
  )
    return undefined;
  const definition = tableStyleDefinition(op.styleId);
  return definition ? [definition] : undefined;
}
