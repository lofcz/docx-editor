import { expect, test } from 'bun:test';
import { serializeLinkedMention } from './mention-literal';
import { parseTargetMarkdown, type TargetAstNode } from './parse-target';
function mentions(node: TargetAstNode): TargetAstNode[] {
  return [
    ...(node.type === 'mention' ? [node] : []),
    ...((node.children ?? []) as TargetAstNode[]).flatMap(mentions),
  ];
}
for (const [key, value] of [
  ['user-7', 'Ada'],
  ['user&α"', 'A "name" [x] | <>&\\'],
  ['user-empty', ''],
])
  test(`literal linked mention retains exact stored values ${JSON.stringify([key, value])}`, () => {
    const serialized = serializeLinkedMention(key!, value!);
    for (const source of [
      `[${serialized.text}](https://outer.test "Read")`,
      `| H |\n| - |\n| [${serialized.text}](https://outer.test "Read") |`,
    ]) {
      const parsed = parseTargetMarkdown(source);
      if (!parsed.success) throw Error(parsed.error);
      const recovered = mentions(parsed.target.normalizedAst);
      expect(recovered.length).toBe(1);
      expect(recovered[0]!.username).toBe(key);
      expect(recovered[0]!.displayText).toBe(value);
      expect(recovered[0]!.children).toEqual([{ type: 'text', value }]);
    }
  });
for (const source of [
  '<scio-mention key="user" value="Ada" extra="keep" />',
  '<scio-mention key="user" key="other" value="Ada" />',
  '<scio-mention key={fetch("asset:x")} value="Ada" />',
  '<scio-mention key="user" value="Ada">keep</scio-mention>',
])
  test(`unknown literal mention data stays opaque: ${source}`, () => {
    const parsed = parseTargetMarkdown(`[${source}](https://outer.test)`);
    if (!parsed.success) throw Error(parsed.error);
    expect(mentions(parsed.target.normalizedAst)).toEqual([]);
    expect(JSON.stringify(parsed.target.normalizedAst)).toContain('scio-mention');
  });
