import { type DocumentNode } from './document-types';

function headingText(node: DocumentNode): string {
  if (node.type === 'image' || node.type === 'imageReference') {
    return typeof node.alt === 'string' ? node.alt : '';
  }
  if (node.type === 'break') return ' ';
  return node.value ?? node.children?.map(headingText).join('') ?? '';
}

/** Assign a unique Markdown fragment to each heading, in story order. */
export class HeadingFragments {
  private readonly used = new Set<string>();
  private readonly suffixes = new Map<string, number>();

  next(node: DocumentNode): string {
    const base = headingText(node)
      .trim()
      .toLowerCase()
      .replace(/[^\p{L}\p{M}\p{N}_\s-]/gu, '')
      .replace(/\s/g, '-');
    let fragment = base;
    let suffix = this.suffixes.get(base) ?? 0;
    while (this.used.has(fragment)) fragment = base + '-' + ++suffix;
    this.suffixes.set(base, suffix);
    this.used.add(fragment);
    return fragment;
  }
}

export function decodeFragment(fragment: string): string {
  try {
    return decodeURIComponent(fragment);
  } catch {
    return fragment;
  }
}
