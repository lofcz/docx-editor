/** Word property snapshots do not own paragraph marks, section state, or break revisions. */
export function independentPropertyChild(parentName: string, childName: string): boolean {
  if (parentName === 'pPr') return childName === 'rPr' || childName === 'sectPr';
  if (parentName === 'trPr') return childName === 'ins' || childName === 'del';
  if (parentName === 'tcPr') return ['cellIns', 'cellDel', 'cellMerge'].includes(childName);
  return parentName === 'rPr' && ['ins', 'del', 'moveFrom', 'moveTo'].includes(childName);
}
