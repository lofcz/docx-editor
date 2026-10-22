import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import fixture from '../__tests__/fixtures/target-validation.json';
import tools from '../__tests__/fixtures/tool-contracts.json';
import { docxFixture } from '../../test/support/document';
import { directParagraphTextOps, planParagraphText } from '../diff/paragraph-text';
import { validateTargetMarkdown, type TargetValidationOptions } from './validate-target';
import { analyzeEditTagBalance } from './validation/tag-balance';

test('the validation oracle records independent application sources and both diagnostic shapes', () => {
  expect(fixture.schemaVersion).toBe(1);
  expect(fixture.source.revision).toMatch(/^[a-f0-9]{40}$/);
  expect(Object.keys(fixture.source.files)).toHaveLength(7);
  for (const hash of Object.values(fixture.source.files)) expect(hash).toMatch(/^[a-f0-9]{64}$/);
  expect(fixture.cases).toHaveLength(32);
  expect(fixture.tagCases).toHaveLength(9);
  const expression = fixture.cases.find((item) => item.name === 'invalid expression')!;
  expect(expression.syntax.issues).toEqual([
    { code: 'acorn', message: 'Could not parse expression with acorn' },
  ]);
  expect(expression.locatedSyntax.issues).toEqual([
    { code: 'acorn', message: 'Could not parse expression with acorn', line: 1, column: 14 },
  ]);
});

for (const item of fixture.cases)
  test(`application syntax, schema, and exact feedback: ${item.name}`, () => {
    for (const operation of ['edit', 'set'] as const) {
      const result = validateTargetMarkdown(item.input, { materialKind: 'lesson_plan', operation });
      if (!item.syntax.ok || (!item.lesson.ok && operation === 'set')) {
        expect(result.success).toBe(false);
        if (result.success) throw new Error('Expected rejection');
        expect(result.stage).toBe(item.syntax.ok ? 'schema' : 'syntax');
        expect(result.issues).toEqual(
          item.syntax.ok ? item.lesson.issues : item.locatedSyntax.issues
        );
        expect(result.error).toBe(item.locatedRejection!);
        expect(Object.hasOwn(result, 'target')).toBe(false);
      } else {
        expect(result.success).toBe(true);
        if (!result.success) throw new Error(result.error);
        expect(result.warnings).toEqual(item.lesson.issues);
        expect(result.warning).toBe('warning' in item ? item.warning : undefined);
        expect(result.target.sourceMarkdown).toBe(item.input);
        expect(Object.isFrozen(result.target.ast)).toBe(true);
      }
      expect(Object.isFrozen(result)).toBe(true);
      expect(Object.isFrozen(result.success ? result.warnings : result.issues)).toBe(true);
    }
    for (const materialKind of ['document', 'curriculum'] as const) {
      const result = validateTargetMarkdown(item.input, { materialKind, operation: 'set' });
      expect(result.success).toBe(item.syntax.ok);
      if (result.success) {
        expect(result.warnings).toEqual([]);
        expect(result.warning).toBeUndefined();
      }
    }
  });

test.each(fixture.tagCases)('application tag-balance hints: $oldString -> $newString', (item) => {
  expect(analyzeEditTagBalance(item.oldString, item.newString)).toEqual(item.hints);
});

test('C02 malformed edit and full-set errors retain exact tool text, hints, and rejected context', () => {
  for (const name of ['malformed edit', 'set invalid syntax', 'set invalid lesson schema']) {
    const item = tools.executorCases.find((item) => item.name === name)!;
    const isEdit = item.tool === 'edit_document';
    const input = item.input as Record<string, unknown>;
    const result = validateTargetMarkdown(isEdit ? '<activity>Broken\n' : String(input.markdown), {
      materialKind: item.materialKind as TargetValidationOptions['materialKind'],
      operation: isEdit ? 'edit' : 'set',
      ...(isEdit
        ? {
            replacement: {
              oldString: String(input.old_string),
              newString: String(input.new_string),
            },
          }
        : {}),
    });
    expect(result.success).toBe(false);
    if (result.success) throw new Error('Expected rejection');
    expect(result.error).toBe(item.expected.result.error!);
  }
  const item = tools.executorCases.find((item) => item.name === 'incremental shape warning')!;
  const result = validateTargetMarkdown('Beta.\n\n', {
    materialKind: 'lesson_plan',
    operation: 'edit',
  });
  expect(result.success).toBe(true);
  if (!result.success) throw new Error(result.error);
  expect(item.expected.result.message).toContain(result.warning!);
});

test('schema warnings retain the target; rejection prevents any native mutation', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docxFixture('<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>'),
  });
  try {
    const surface = editor.surface!,
      session = surface.session,
      part = session.part(),
      id = session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 2 },
      head: { paragraphId: id, offset: 2 },
    });
    const state = {
      fingerprint: canonicalOoxmlFingerprint(part),
      revision: session.packageRevision(),
      selection: structuredClone(surface.state().selection),
      history: [session.canUndo(), session.canRedo()],
    };
    for (const [input, operation] of [
      ['<activity>Broken', 'edit'],
      ['Beta.', 'set'],
    ] as const) {
      const result = validateTargetMarkdown(input, { materialKind: 'lesson_plan', operation });
      expect(result.success).toBe(false);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(state.fingerprint);
      expect(session.packageRevision()).toBe(state.revision);
      expect(surface.state().selection).toEqual(state.selection);
      expect([session.canUndo(), session.canRedo()]).toEqual(state.history);
    }
    const allowed = validateTargetMarkdown('Beta.', {
      materialKind: 'lesson_plan',
      operation: 'edit',
    });
    if (!allowed.success) throw new Error(allowed.error);
    expect(allowed.warnings.map((issue) => issue.code)).toEqual([
      'missing-lesson-title',
      'missing-preparation-recap',
      'phase-count',
    ]);
    expect(allowed.target.markdown).toBe('Beta.');
    const paragraph = allowed.target.normalizedAst.children[0]!;
    const text = (paragraph.children as { type: string; value: string }[])
      .map((child) => child.value)
      .join('');
    const planned = planParagraphText(
      {
        revision: { documentId: 'validation', epoch: 'mount', revision: state.revision },
        partName: part.name,
        paragraphId: id,
        text: 'Alpha.',
      },
      text
    );
    if (!planned.success) throw new Error(planned.error);
    let after = part;
    for (const edit of planned.plan.operations) {
      if (edit.kind !== 'text') throw new Error('Expected paragraph text');
      for (const op of directParagraphTextOps(edit)) {
        const applied = applyTreeOp(after, op);
        if (!applied.ok) throw new Error(applied.reason);
        after = applied.part;
      }
    }
    expect(paragraphTextOf(after, id)).toBe('Beta.');
    expect(canonicalOoxmlFingerprint(session.part())).toBe(state.fingerprint);
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('expressions stay inert and rejected diagnostics use normalized target line positions', () => {
  const env = globalThis as typeof globalThis & { __validationProbe?: () => void };
  let calls = 0;
  env.__validationProbe = () => {
    calls++;
  };
  try {
    const result = validateTargetMarkdown(
      '<future expr={globalThis.__validationProbe()}>\n\n{globalThis.__validationProbe()}\n\n</future>',
      {
        materialKind: 'document',
        operation: 'set',
      }
    );
    expect(result.success).toBe(true);
    expect(calls).toBe(0);
    const bad = validateTargetMarkdown('Label:\n6. First\n7. Second\n\n</bad>', {
      materialKind: 'document',
      operation: 'edit',
    });
    expect(bad.success).toBe(false);
    if (bad.success) throw new Error('Expected rejection');
    expect(bad.issues[0]!.line).toBe(6);
    expect(bad.error).toContain('>>>    6 | </bad>');
  } finally {
    delete env.__validationProbe;
  }
});

test('invalid options, accessors, excessive input, and parse limits reject before returning a target', () => {
  let getters = 0;
  const accessor = Object.defineProperty({}, 'materialKind', {
    enumerable: true,
    get() {
      getters++;
      return 'document';
    },
  });
  for (const options of [
    null,
    [],
    accessor,
    { materialKind: 'unknown', operation: 'set' },
    { materialKind: 'document', operation: 'bad' },
    { materialKind: 'document', operation: 'set', extra: true },
    { materialKind: 'document', operation: 'set', replacement: { oldString: 'A' } },
    { materialKind: 'document', operation: 'set', limits: { maxBytes: 0 } },
    { materialKind: 'document', operation: 'set', limits: { unknown: 2 } },
  ]) {
    const result = validateTargetMarkdown('Text.', options as TargetValidationOptions);
    expect(result.success).toBe(false);
    if (result.success) throw new Error('Expected rejection');
    expect(result.stage).toBe('input');
    expect(Object.hasOwn(result, 'target')).toBe(false);
  }
  expect(getters).toBe(0);
  for (const input of [42, 'x'.repeat(2 * 1024 * 1024 + 1)]) {
    const result = validateTargetMarkdown(input as string, {
      materialKind: 'document',
      operation: 'edit',
    });
    expect(result.success).toBe(false);
  }
  const limited = validateTargetMarkdown('<a><b>Text</b></a>', {
    materialKind: 'document',
    operation: 'set',
    limits: { maxDepth: 3 },
  });
  expect(limited.success).toBe(false);
  if (limited.success) throw new Error('Expected rejection');
  expect(limited.issues[0]!.code).toBe('depth-limit');
});

test('partial malformed tag scanning stays bounded and limits duplicate feedback', () => {
  expect(analyzeEditTagBalance('<a '.repeat(50_000), '<callout>')).toHaveLength(1);
  const hints = analyzeEditTagBalance('<activity>'.repeat(70), '');
  expect(hints).toHaveLength(65);
  expect(hints[0]).toContain('Your old_string contains <activity>');
  expect(hints[63]).toBe(hints[0]);
  expect(hints[64]).toBe('Additional tag-boundary hints were omitted after the first 64.');
  expect(analyzeEditTagBalance('<a><b></a>', '<b>')).toEqual([]);
});
