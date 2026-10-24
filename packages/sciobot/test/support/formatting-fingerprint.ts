import { canonicalOoxmlFingerprint, type OoxmlPart } from '@docx-editor.dev/core/store';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML = 'http://www.w3.org/XML/1998/namespace';
type Token = unknown[];
const element = (value: unknown, name: string): value is Token =>
  Array.isArray(value) && value[0] === 'element' && value[1] === W && value[2] === name;
const equal = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);

/**
 * Word formatting can split a text-only run without changing its meaning. Compare the
 * complete canonical native tree while allowing only adjacent runs with identical full
 * attributes and property subtrees to share one text token. Fields, drawings, controls,
 * foreign children, and every other structural difference remain significant.
 */
export function formattingFingerprint(part: OoxmlPart): string {
  const plainRun = (run: unknown): run is Token => {
    if (!element(run, 'r')) return false;
    const children = run[4] as Token[];
    return (
      (children.length === 1 || (children.length === 2 && element(children[0], 'rPr'))) &&
      element(children.at(-1), 't') &&
      (children.at(-1)![3] as Token[]).every((a) => equal(a.slice(0, 2), [XML, 'space'])) &&
      (children.at(-1)![4] as Token[]).every((t) => t[0] === 'text')
    );
  };
  const text = (run: Token) =>
    ((run[4] as Token[]).at(-1)![4] as Token[]).map((t) => t[1]).join('');
  const properties = (run: Token) => (run[4] as Token[]).filter((child) => !element(child, 't'));
  const walk = (token: Token): Token => {
    if (token[0] !== 'element') return token;
    const children: Token[] = [];
    for (const raw of token[4] as Token[]) {
      const child = walk(raw),
        previous = children.at(-1);
      if (
        plainRun(previous) &&
        plainRun(child) &&
        equal(previous[3], child[3]) &&
        equal(properties(previous), properties(child))
      ) {
        const joined = text(previous) + text(child);
        const attributes = /^\s|\s$/.test(joined) ? [[XML, 'space', 'preserve']] : [];
        previous[4] = [
          ...properties(previous),
          ['element', W, 't', attributes, [['text', joined]]],
        ];
      } else children.push(child);
    }
    return [...token.slice(0, 4), children];
  };
  return JSON.stringify(walk(JSON.parse(canonicalOoxmlFingerprint(part))));
}
