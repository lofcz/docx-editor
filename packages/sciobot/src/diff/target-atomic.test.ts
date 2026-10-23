import { parseTargetMarkdown } from '../projection/parse-target';
import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { atomicFixture, atomicListening, atomicVisual } from '../../test/support/atomic-execution';
import { propertyEditor, propertyToken } from '../../test/support/property-execution';
import type { BlockPayload } from '../blocks/payload-schema';
import type { JsonObject, JsonValue } from '../data/json';
import { readWorkingView } from '../review/working-view';
import { planTargetDocument } from './plan-target';

function patchText(value: JsonValue, before: string, after: string): JsonValue {
  if (Array.isArray(value)) return value.map((item) => patchText(item, before, after));
  if (value && typeof value === 'object')
    return Object.fromEntries(
      Object.entries(value).map(([key, child]) => [
        key,
        key === 'value' && child === before ? after : patchText(child, before, after),
      ])
    );
  return value;
}
for (const kind of ['visual', 'listening', 'custom'] as const)
  test(`complete ${kind} target retains the full atomic payload and unprojected data`, async () => {
    let payload: BlockPayload =
      kind === 'visual'
        ? {
            ...atomicVisual,
            attributes: { ...atomicVisual.attributes, _private: { keep: true }, absentNull: null },
          }
        : structuredClone(atomicListening);
    if (kind !== 'visual') {
      const mdast = payload.mdast as JsonObject;
      const children = mdast.children as readonly JsonObject[];
      payload = {
        ...payload,
        mdast: {
          ...mdast,
          retainedRootData: { future: ['keep', 7] },
          children: children.map((child, index) =>
            index === 0
              ? { ...child, retainedChildData: { portrait: 'keep-portrait', voice: 'keep-voice' } }
              : child
          ),
          ...(kind === 'custom' ? { name: 'NovelReactPanel' } : {}),
        },
      };
    }
    const editor = propertyEditor(atomicFixture([payload]));
    try {
      const session = editor.surface!.session;
      const before = session.currentPackage();
      const fingerprint = canonicalOoxmlFingerprint(session.part());
      const view = await readWorkingView(session, propertyToken(editor));
      if (!view.success) throw Error(view.error);
      const target =
        kind === 'visual'
          ? view.view.projection.text
              .replace('vfs://pristine.png', 'vfs://updated.png')
              .replace('alt="Old picture"', 'alt="Updated picture"')
          : view.view.projection.text
              .replace('original-audio', 'updated-audio')
              .replace('Good morning!', 'Good night!');
      const planned = planTargetDocument(
        { pkg: before, revision: propertyToken(editor), projection: view.view.projection },
        target
      );
      if (!planned.success) throw Error(planned.error);
      expect(planned.plan.operations).toHaveLength(1);
      const edit = planned.plan.operations[0]!;
      expect(edit.kind).toBe('payload');
      if (edit.kind !== 'payload') throw Error('Wrong atomic operation');
      expect(edit.path).toEqual([]);
      expect(edit.before).toEqual({ present: true, value: payload });
      const expected =
        kind === 'visual'
          ? {
              ...payload,
              attributes: {
                ...payload.attributes,
                src: 'vfs://updated.png',
                alt: 'Updated picture',
              },
            }
          : {
              ...payload,
              mdast: patchText(
                patchText(payload.mdast!, 'original-audio', 'updated-audio'),
                'Good morning!',
                'Good night!'
              ),
            };
      expect(edit.after).toEqual({ present: true, value: expected });
      expect(session.currentPackage()).toBe(before);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(fingerprint);
      expect(session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

for (const action of ['reorder', 'remove', 'rename'] as const)
  test(`opaque ${action} keeps private child data with its explicit syntax identity`, async () => {
    const parsed = parseTargetMarkdown(
      '<NovelReactPanel>\n\n<Widget id="a">Alpha</Widget>\n\n<Widget id="b">Beta</Widget>\n\n</NovelReactPanel>'
    );
    if (!parsed.success) throw Error(parsed.error);
    const raw = parsed.target.normalizedAst.children[0] as JsonObject;
    const children = (raw.children as readonly JsonObject[]).map((child, index) => ({
      ...child,
      privateData: { owner: index === 0 ? 'A' : 'B', file: index === 0 ? 'keep-A' : 'keep-B' },
    }));
    const mdast: JsonObject = { ...raw, children, privateData: { keep: 'root' } };
    const payload = { ...atomicListening, mdast };
    const editor = propertyEditor(atomicFixture([payload]));
    try {
      const view = await readWorkingView(editor.surface!.session, propertyToken(editor));
      if (!view.success) throw Error(view.error);
      expect(view.view.projection.text).toContain('<Widget id="a">Alpha</Widget>');
      const target =
        action === 'rename'
          ? view.view.projection.text.replaceAll('NovelReactPanel', 'RenamedReactPanel')
          : action === 'reorder'
            ? view.view.projection.text
                .replace('<Widget id="a">Alpha</Widget>', '<Widget id="b">Beta revised</Widget>')
                .replace('<Widget id="b">Beta</Widget>', '<Widget id="a">Alpha revised</Widget>')
            : view.view.projection.text
                .replace('<Widget id="a">Alpha</Widget>', '')
                .replace('Beta', 'Beta revised');
      const planned = planTargetDocument(
        {
          pkg: editor.surface!.session.currentPackage(),
          revision: propertyToken(editor),
          projection: view.view.projection,
        },
        target
      );
      if (!planned.success) throw Error(planned.error);
      expect(planned.plan.operations).toHaveLength(1);
      const edit = planned.plan.operations[0]!;
      if (edit.kind !== 'payload') throw Error('Missing atomic operation');
      const expectedChildren =
        action === 'rename'
          ? children
          : action === 'reorder'
            ? [
                patchText(children[1]!, 'Beta', 'Beta revised'),
                patchText(children[0]!, 'Alpha', 'Alpha revised'),
              ]
            : [patchText(children[1]!, 'Beta', 'Beta revised')];
      expect(edit.after).toEqual({
        present: true,
        value: {
          ...payload,
          mdast: {
            ...mdast,
            children: expectedChildren,
            ...(action === 'rename' ? { name: 'RenamedReactPanel' } : {}),
          },
        },
      });
    } finally {
      editor.destroy();
    }
  });
