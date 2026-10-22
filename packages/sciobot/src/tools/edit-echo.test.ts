import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import fixture from '../__tests__/fixtures/edit-echo.json';
import { docxFixture, W } from '../../test/support/document';
import type { RevisionToken } from '../contracts/document';
import { directParagraphTextOps, planParagraphText } from '../diff/paragraph-text';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { resolvePatchSpan } from '../projection/resolve-patch';
import { parseTargetMarkdown } from '../projection/parse-target';
import { findAndReplace } from './find-replace';
import { validateTargetMarkdown } from './validate-target';
import { buildEditEcho, type EditEchoProjection, EDIT_ECHO_MAX_LINES } from './edit-echo';

const revision = (value: number): RevisionToken => ({
  documentId: 'echo-document',
  epoch: 'mount-1',
  revision: value,
});
const projection = (text: string, value: number): EditEchoProjection => ({
  text,
  revision: revision(value),
  displayMode: 'proposed',
});

for (const item of fixture.cases)
  test(`application landed excerpt: ${item.name}`, () => {
    const before = projection(item.before, 1),
      landed = projection(item.after, 2);
    const snapshots = JSON.stringify([before, landed]);
    const result = buildEditEcho(before, landed, revision(2));
    if (!result.success) throw new Error(result.error);
    expect(result.echo.message).toBe(item.message);
    expect(result.echo.excerpt).toBe(
      item.after
        .split('\n')
        .slice(result.echo.startLine - 1, result.echo.endLine)
        .join('\n')
    );
    expect(result.echo.endLine - result.echo.startLine + 1).toBeLessThanOrEqual(
      EDIT_ECHO_MAX_LINES
    );
    expect(result.echo.message.includes('excerpt truncated')).toBe(result.echo.truncated);
    expect(JSON.stringify([before, landed])).toBe(snapshots);
    expect(Object.isFrozen(result.echo)).toBe(true);
    expect(Object.isFrozen(result.echo.revision)).toBe(true);
    if (result.echo.excerpt)
      expect(findAndReplace(item.after, result.echo.excerpt, 'Next', false).success).toBe(true);
    expect(Reflect.set(landed, 'revision', revision(99))).toBe(true);
    expect(result.echo.revision).toEqual(revision(2));
  });

test('echoes reject stale, foreign, future, invalid, and inconsistent snapshots without returning excerpts', () => {
  const before = projection('Before', 1),
    landed = projection('After', 2);
  for (const [old, next, actual, code] of [
    [before, landed, revision(3), 'stale-revision'],
    [
      { ...before, revision: { ...revision(1), documentId: 'other' } },
      landed,
      revision(2),
      'stale-revision',
    ],
    [
      { ...before, revision: { ...revision(1), epoch: 'other' } },
      landed,
      revision(2),
      'stale-revision',
    ],
    [projection('Before', 3), landed, revision(2), 'stale-revision'],
    [projection('Before', 2), landed, revision(2), 'invalid-document'],
    [{ ...before, displayMode: 'original' }, landed, revision(2), 'unsupported-operation'],
    [before, { ...landed, displayMode: 'allMarkup' }, revision(2), 'unsupported-operation'],
    ['Raw model text', landed, revision(2), 'invalid-input'],
    [before, null, revision(2), 'invalid-input'],
    [before, landed, { ...revision(2), revision: -1 }, 'invalid-input'],
  ] as const) {
    const result = buildEditEcho(old as EditEchoProjection, next as EditEchoProjection, actual);
    expect(result.success).toBe(false);
    if (result.success) throw new Error('Expected refusal');
    expect(result.code).toBe(code);
    expect(Object.hasOwn(result, 'echo')).toBe(false);
  }
  let calls = 0;
  const accessor = Object.defineProperty({ ...landed }, 'text', {
    enumerable: true,
    get() {
      calls++;
      return 'Unexpected';
    },
  });
  expect(buildEditEcho(before, accessor, revision(2)).success).toBe(false);
  expect(calls).toBe(0);
  const noChange = buildEditEcho(before, before, revision(1));
  expect(noChange.success).toBe(true);
});

test('a copied landed echo chains a second native patch after list ordinals and whitespace normalize', async () => {
  const p = (text: string, list = false) =>
    '<w:p>' +
    (list ? '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' : '') +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docxFixture(p('Lead') + p('First', true) + p('Second', true) + p('Tail'), {
      numbering,
    }),
  });
  const surface = editor.surface!,
    session = surface.session;
  const token = () => readEditorRevision(editor, 'echo-document', 'mount-1')!;
  const working = async () => {
    const result = await readWorkingView(session, token());
    if (!result.success) throw new Error(result.error);
    return result.view.projection;
  };
  const apply = (id: string, target: string) => {
    const part = session.part(),
      planned = planParagraphText(
        {
          revision: token(),
          partName: part.name,
          paragraphId: id,
          text: paragraphTextOf(part, id)!,
        },
        target
      );
    if (!planned.success) throw new Error(planned.error);
    const result = surface.applyAutomationOps(() =>
      planned.plan.operations.flatMap((edit) =>
        edit.kind === 'text' ? directParagraphTextOps(edit) : []
      )
    );
    expect(result.rejected).toBe(false);
  };
  const parsed = (text: string) => {
    const result = parseTargetMarkdown(text);
    if (!result.success) throw new Error(result.error);
    return result.target.normalizedAst;
  };
  try {
    const ids = session.paragraphIds(),
      before = await working();
    expect(before.text).toBe('Lead\n\n1. First\n\n2. Second\n\nTail');
    const requested = '\n\nLead\n\n1. Updated item\n\n1. Second\n\nTail\n\n\n';
    const validation = validateTargetMarkdown(requested, {
      materialKind: 'document',
      operation: 'edit',
    });
    if (!validation.success) throw new Error(validation.error);
    apply(ids[1]!, 'Updated item');
    const landed = await working();
    expect(parsed(requested)).toEqual(parsed(landed.text));
    expect(landed.text).not.toBe(requested);
    const fingerprint = canonicalOoxmlFingerprint(session.part()),
      history = [session.canUndo(), session.canRedo()];
    const first = buildEditEcho(before, landed, token());
    if (!first.success) throw new Error(first.error);
    expect(first.echo.message).not.toContain('1. Second');
    const copied = first.echo.excerpt.split('\n').find((line) => line === '2. Second')!;
    expect(copied).toBe('2. Second');
    expect(requested.includes(copied)).toBe(false);
    const patched = findAndReplace(landed.text, copied, '2. Final', false);
    if (!patched.success) throw new Error(patched.error);
    const secondValidation = validateTargetMarkdown(patched.newContent, {
      materialKind: 'document',
      operation: 'edit',
      replacement: { oldString: copied, newString: '2. Final' },
    });
    expect(secondValidation.success).toBe(true);
    const at = landed.text.indexOf(copied),
      resolved = resolvePatchSpan(
        landed,
        { start: at, end: at + copied.length },
        { revision: token() }
      );
    if (!resolved.success || resolved.target.kind !== 'paragraph')
      throw new Error('Expected a native paragraph address');
    expect(resolved.target.paragraphId).toBe(ids[2]!);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(fingerprint);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    apply(resolved.target.paragraphId, 'Final');
    const final = await working();
    expect(final.text).toBe('Lead\n\n1. Updated item\n\n2. Final\n\nTail');
    expect(parsed(patched.newContent)).toEqual(parsed(final.text));
    expect(session.paragraphIds()).toEqual(ids);
    const second = buildEditEcho(landed, final, token());
    if (!second.success) throw new Error(second.error);
    expect(second.echo.excerpt).toContain('2. Final');
    const stale = buildEditEcho(before, landed, token());
    expect(stale.success).toBe(false);
    if (!stale.success) expect(stale.code).toBe('stale-revision');
  } finally {
    editor.destroy();
    container.remove();
  }
});
