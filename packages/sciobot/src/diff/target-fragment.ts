/** The D01 JSON fragment boundary, shared by structural planners and plan decoding. */
export function isTargetFragment(value: unknown): boolean {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const node = value as Record<string, unknown>;
  if (
    typeof node.type !== 'string' ||
    !node.type ||
    node.type === 'root' ||
    Object.hasOwn(node, 'kind')
  )
    return false;
  return (
    !Object.hasOwn(node, 'children') ||
    (Array.isArray(node.children) && node.children.every(isTargetFragment))
  );
}
