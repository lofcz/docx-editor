import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { checklistParagraph, listDocument, listParagraph } from '../../test/support/list-document';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { directParagraphTextOps, planParagraphText } from '../diff/paragraph-text';
import { planListMembership } from '../diff/lists';
import { readWorkingView } from '../review/working-view';
import { readCommittedView } from '../review/committed-view';
import { parseTargetMarkdown, type TargetAstNode } from './parse-target';
import { projectLists } from './lists';
import { markdownPointAt, projectMarkdown } from './markdown';
import { projectMdx } from './mdx';
import { collectProtectedContent } from './protected-content';
import { resolvePatchSpan } from './resolve-patch';

function ast(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  const items: TargetAstNode[] = [];
  const visit = (node: TargetAstNode) => {
    if (node.type === 'listItem') items.push(node);
    if (Array.isArray(node.children)) (node.children as TargetAstNode[]).forEach(visit);
  };
  visit(parsed.target.normalizedAst);
  return items;
}
async function imported(
  body: string,
  mode: 'proposed' | 'original' | 'all-markup' = 'proposed',
  withDomain: boolean | 'lesson_visual' = false
) {
  let bytes = listDocument(body);
  if (withDomain) {
    const parsed = readOoxmlPackage(bytes);
    if (!parsed.ok) throw new Error(parsed.reason);
    const written = withPayloadRecord(
      parsed.package,
      {
        id: 'activity',
        type: withDomain === 'lesson_visual' ? withDomain : 'lesson_activity',
        version: 1,
        attributes: { name: 'Tasks', duration: '5' },
      },
      ''
    );
    if (!written.ok) throw new Error(written.reason);
    bytes = writeOoxmlPackage(written.pkg);
  }
  const editor = createDocxEditor({ container: document.createElement('div'), document: bytes });
  const session = editor.surface!.session,
    opened = openDocumentForExport(session);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const revision = {
      documentId: 'checklists',
      epoch: 'mount-1',
      revision: session.packageRevision(),
    };
    const layout = await opened.session.layoutFor(mode);
    const result = withDomain
      ? projectMdx(session.currentPackage(), layout, revision)
      : projectLists(session.currentPackage(), layout, revision);
    if (!result.success) throw new Error(result.error);
    return { editor, session, revision, layout, projection: result.projection };
  } catch (cause) {
    editor.destroy();
    throw cause;
  } finally {
    opened.session.dispose();
  }
}

test('owned checked and unchecked markers become immutable mapped GFM tasks with nesting', async () => {
  const value = await imported(checklistParagraph(false) + checklistParagraph(true, 1));
  try {
    const { projection, session, revision } = value;
    expect(projection.text).toBe('- [ ] Task\n\n  - [x] Task');
    expect(ast(projection.text).map((node) => node.checked)).toEqual([false, true]);
    expect(projection.checklists.map((item) => item.level)).toEqual([0, 1]);
    for (const item of projection.checklists) {
      expect(projection.text.slice(item.stateStart, item.stateEnd)).toBe(
        item.checked ? '[x]' : '[ ]'
      );
      expect(markdownPointAt(projection, item.contentStart)).toMatchObject({
        paragraphId: item.paragraphId,
        offset: 1,
      });
      expect(markdownPointAt(projection, item.stateStart)).toBeNull();
      expect(
        resolvePatchSpan(projection, { start: item.stateStart, end: item.stateEnd }, { revision })
      ).toMatchObject({
        success: true,
        target: {
          kind: 'checklist-state',
          paragraphId: item.paragraphId,
          controlId: item.controlId,
          checked: item.checked,
          level: item.level,
        },
      });
      expect(
        resolvePatchSpan(
          projection,
          { start: item.stateStart + 1, end: item.stateEnd - 1 },
          { revision }
        )
      ).toMatchObject({
        success: true,
        target: { kind: 'checklist-state', controlId: item.controlId },
      });
      const requested = planListMembership(
        { pkg: session.currentPackage(), partName: session.part().name, revision },
        [
          {
            paragraphId: item.paragraphId,
            target: { kind: 'checklist', checked: !item.checked, level: item.level },
          },
        ]
      );
      expect(requested.success && requested.plan.operations.length).toBe(1);
    }
    expect(
      Object.isFrozen(projection) &&
        Object.isFrozen(projection.checklists) &&
        projection.checklists.every(Object.isFrozen)
    ).toBe(true);
    expect(JSON.parse(JSON.stringify(projection))).toEqual(projection);
    expect(session.canUndo()).toBe(false);
  } finally {
    value.editor.destroy();
  }
});

test('documents with ordinary form controls retain the native Markdown and source map unchanged', async () => {
  const value = await imported(
    checklistParagraph(false, 0, 'ordinary-form') + listParagraph('Other')
  );
  try {
    const native = projectMarkdown(value.layout, value.revision);
    expect(value.projection.checklists).toEqual([]);
    expect(value.projection.text).toBe(native.text);
    expect(value.projection.sourceMap).toEqual(native.sourceMap);
    expect(value.projection.syntax).toEqual(native.syntax);
  } finally {
    value.editor.destroy();
  }
});

test('native text edits retain checkbox identity and state through undo/redo and save/reopen', async () => {
  const value = await imported(checklistParagraph(false));
  try {
    const { projection, session, revision, editor } = value,
      item = projection.checklists[0]!;
    const node = findNode(session.part(), item.controlId)!,
      before = canonicalOoxmlFingerprint(node);
    const protectedResult = collectProtectedContent(session.currentPackage(), projection, revision);
    if (!protectedResult.success) throw new Error(protectedResult.error);
    const start = projection.text.indexOf('Task'),
      resolved = resolvePatchSpan(
        projection,
        { start, end: start + 4 },
        { revision, protected: protectedResult.index.boundaries }
      );
    expect(resolved).toMatchObject({
      success: true,
      target: { kind: 'text', start: { offset: 1 }, end: { offset: 5 } },
    });
    const planned = planParagraphText(
      {
        revision,
        partName: session.part().name,
        paragraphId: item.paragraphId,
        offset: 1,
        text: 'Task',
      },
      'Edited 😀 task'
    );
    if (!planned.success) throw new Error(planned.error);
    const ops = planned.plan.operations.flatMap((edit) =>
      edit.kind === 'text' ? directParagraphTextOps(edit) : []
    );
    expect(editor.surface!.applyAutomationOps(() => ops).rejected).toBe(false);
    expect(findNode(session.part(), item.controlId) === node).toBe(true);
    expect(canonicalOoxmlFingerprint(findNode(session.part(), item.controlId)!)).toBe(before);
    const current = { ...revision, revision: session.packageRevision() };
    const working = await readWorkingView(session, current);
    if (!working.success) throw new Error(working.error);
    expect(working.view.projection.text).toBe('- [ ] Edited 😀 task');
    session.undo();
    expect(paragraphTextOf(session.part(), item.paragraphId)).toBe('☐Task');
    session.redo();
    expect(paragraphTextOf(session.part(), item.paragraphId)).toBe('☐Edited 😀 task');
    const reopened = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(session.currentPackage()),
    });
    try {
      const next = await readWorkingView(reopened.surface!.session, {
        ...revision,
        revision: reopened.surface!.session.packageRevision(),
      });
      if (!next.success) throw new Error(next.error);
      expect(next.view.projection.text).toBe('- [ ] Edited 😀 task');
      expect(next.view.projection.checklists.map((item) => item.checked)).toEqual([false]);
    } finally {
      reopened.destroy();
    }
  } finally {
    value.editor.destroy();
  }
});

test('a resolved state-only patch changes its native control while retaining task text and source positions', async () => {
  const value = await imported(checklistParagraph(false));
  try {
    const { projection, revision, session, editor } = value,
      task = projection.checklists[0]!;
    const resolved = resolvePatchSpan(
      projection,
      { start: task.stateStart, end: task.stateEnd },
      { revision }
    );
    if (!resolved.success || resolved.target.kind !== 'checklist-state')
      throw new Error('Missing task state target');
    expect(
      editor.surface!.applyAutomationOps(() => [
        {
          op: 'setContentControlValue',
          controlId: resolved.target.kind === 'checklist-state' ? resolved.target.controlId : '',
          value: { kind: 'checkbox', checked: true },
        },
      ]).rejected
    ).toBe(false);
    const next = await readWorkingView(session, {
      ...revision,
      revision: session.packageRevision(),
    });
    if (!next.success) throw new Error(next.error);
    expect(next.view.projection.text).toBe('- [x] Task');
    expect(next.view.projection.checklists[0]!.controlId).toBe(task.controlId);
    expect(
      markdownPointAt(next.view.projection, next.view.projection.checklists[0]!.contentStart)
    ).toMatchObject({ paragraphId: task.paragraphId, offset: 1 });
    const reopened = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(session.currentPackage()),
    });
    try {
      const read = await readWorkingView(reopened.surface!.session, {
        ...revision,
        revision: reopened.surface!.session.packageRevision(),
      });
      if (!read.success) throw new Error(read.error);
      expect(read.view.projection.text).toBe('- [x] Task');
    } finally {
      reopened.destroy();
    }
  } finally {
    value.editor.destroy();
  }
});

test('manual breaks, escapes, marks, inline symbols, and empty tasks keep original native offsets', async () => {
  const body =
    checklistParagraph(false).replace(
      '<w:t>Task</w:t>',
      '<w:rPr><w:b/></w:rPr><w:t>A&amp;😀</w:t><w:br/><w:sym w:font="MS Gothic" w:char="2612"/><w:t>tail</w:t>'
    ) + checklistParagraph(true).replace('<w:t>Task</w:t>', '<w:t></w:t>');
  const value = await imported(body);
  try {
    const { projection, revision } = value;
    expect(projection.text).toContain('- [ ] **A&amp;😀**  \n  **☒tail**\n\n- [x] \u00a0');
    expect(projection.text.endsWith('- [x] \u00a0')).toBe(true);
    const at = projection.text.indexOf('tail');
    expect(markdownPointAt(projection, at)).toMatchObject({ offset: 7 });
    expect(ast(projection.text).map((node) => node.checked)).toEqual([false, true]);
    const empty = projection.checklists[1]!;
    expect(markdownPointAt(projection, empty.contentStart)).toMatchObject({
      paragraphId: empty.paragraphId,
      offset: 1,
    });
    expect(
      resolvePatchSpan(projection, { start: empty.stateStart, end: empty.stateEnd }, { revision })
    ).toMatchObject({ success: true, target: { kind: 'checklist-state', checked: true } });
  } finally {
    value.editor.destroy();
  }
});

test('cell tasks use HTML boundaries while ordinary imported form controls retain their native output', async () => {
  const body =
    '<w:tbl><w:tr><w:tc>' +
    checklistParagraph(true) +
    '</w:tc><w:tc>' +
    checklistParagraph(false, 0, 'ordinary-form') +
    '</w:tc></w:tr></w:tbl>';
  const value = await imported(body);
  try {
    const { projection, revision } = value;
    expect(projection.text).toContain('<table>');
    expect(projection.text).toContain('- [x] Task');
    expect(projection.text).toContain('☐Task');
    expect(projection.checklists).toHaveLength(1);
    expect(ast(projection.text).map((node) => node.checked)).toEqual([true]);
    const task = projection.checklists[0]!;
    expect(markdownPointAt(projection, task.contentStart)).toMatchObject({
      paragraphId: task.paragraphId,
      offset: 1,
    });
    expect(
      resolvePatchSpan(projection, { start: task.stateStart, end: task.stateEnd }, { revision })
    ).toMatchObject({ success: true, target: { kind: 'checklist-state' } });
  } finally {
    value.editor.destroy();
  }
});

test('domain composition remaps task state and content positions without exposing atomic fallback tasks', async () => {
  const control = (id: string, content: string) =>
    `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const value = await imported(control('activity', checklistParagraph(false)), 'proposed', true);
  try {
    const { projection, revision } = value,
      item = projection.checklists[0]!;
    expect(projection.text).toContain('- [ ] Task');
    expect(projection.text.slice(item.stateStart, item.stateEnd)).toBe('[ ]');
    expect(markdownPointAt(projection, item.contentStart)).toMatchObject({
      paragraphId: item.paragraphId,
      offset: 1,
    });
    expect(
      resolvePatchSpan(projection, { start: item.stateStart, end: item.stateEnd }, { revision })
    ).toMatchObject({
      success: true,
      target: { kind: 'checklist-state', controlId: item.controlId },
    });
  } finally {
    value.editor.destroy();
  }
  const atomic = await imported(
    control('activity', checklistParagraph(true)),
    'proposed',
    'lesson_visual'
  );
  try {
    expect(atomic.projection.checklists).toEqual([]);
    expect(atomic.projection.text).toContain('<lesson_visual');
    expect(atomic.projection.text).not.toContain('- [x]');
  } finally {
    atomic.editor.destroy();
  }
});

test('empty tasks at document edges and between paragraphs retain distinct constructs and insertion points', async () => {
  const empty = checklistParagraph(false).replace('<w:t>Task</w:t>', '<w:t/>');
  const value = await imported(
    empty + checklistParagraph(true) + empty + listParagraph('Other') + empty
  );
  try {
    const { projection, session } = value;
    expect(projection.text).toBe(
      '- [ ] \u00a0\n\n- [x] Task\n\n- [ ] \u00a0\n\nOther\n\n- [ ] \u00a0'
    );
    expect(ast(projection.text).map((node) => node.checked)).toEqual([false, true, false, false]);
    for (const task of projection.checklists) {
      const bounds = projection.sourceMap.paragraphs.find(
        (p) => p.paragraphId === task.paragraphId
      )!;
      expect([bounds.markdownStart, bounds.markdownEnd]).toEqual([task.start, task.end]);
      expect(markdownPointAt(projection, task.contentStart)).toMatchObject({
        paragraphId: task.paragraphId,
        offset: 1,
      });
    }
    const normal = projection.sourceMap.paragraphs.find(
      (p) => p.paragraphId === session.paragraphIds()[3]
    )!;
    expect(projection.text.slice(normal.markdownStart, normal.markdownEnd)).toBe('Other');
  } finally {
    value.editor.destroy();
  }
});

test('generated task properties use their own ownership while P12 still protects native body boundaries', async () => {
  const value = await imported(
    checklistParagraph(false).replace('<w:sdt>', '<w:bookmarkStart w:id="9" w:name="Keep"/><w:sdt>')
  );
  try {
    const { projection, revision, session } = value,
      task = projection.checklists[0]!;
    const collected = collectProtectedContent(session.currentPackage(), projection, revision);
    if (!collected.success) throw new Error(collected.error);
    const state = { start: task.stateStart, end: task.stateEnd };
    expect(collected.index.boundaries.length > 0).toBe(true);
    expect(
      resolvePatchSpan(projection, state, { revision, protected: collected.index.boundaries })
    ).toMatchObject({ success: true, target: { kind: 'checklist-state' } });
    expect(
      resolvePatchSpan(projection, state, {
        revision,
        protected: [{ ...state, reason: 'Explicit read-only region', scope: 'all' }],
      })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
  } finally {
    value.editor.destroy();
  }
});

test('data-bound and malformed owned controls refuse state writes without changing canonical content', async () => {
  const bound = await imported(
    checklistParagraph(false).replace(
      '<w:sdtPr>',
      '<w:sdtPr><w:dataBinding w:xpath="/x" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/>'
    )
  );
  try {
    const { projection, revision } = bound,
      task = projection.checklists[0]!;
    expect(task.writable).toBe(false);
    expect(
      resolvePatchSpan(projection, { start: task.stateStart, end: task.stateEnd }, { revision })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
  } finally {
    bound.editor.destroy();
  }
  const bytes = listDocument(checklistParagraph(false).replaceAll('w14:checkbox', 'w14:unknown'));
  const loaded = readOoxmlPackage(bytes),
    opened = openDocumentForExport(bytes);
  if (!loaded.ok || !opened.ok) throw new Error('Invalid fixture');
  try {
    const before = canonicalOoxmlFingerprint(
      loaded.package.parts.get(loaded.package.mainDocumentPart)!
    );
    expect(
      projectLists(loaded.package, await opened.session.layout(), {
        documentId: 'bad',
        epoch: 'e',
        revision: 0,
      })
    ).toMatchObject({ success: false, code: 'invalid-document' });
    expect(
      canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
    ).toBe(before);
  } finally {
    opened.session.dispose();
  }
});

test('Original and Proposed select the correct native control with stable model text addresses', async () => {
  const marker = (checked: boolean) => {
    const p = checklistParagraph(checked);
    return p.slice(p.indexOf('<w:sdt>'), p.indexOf('</w:sdt>') + 8);
  };
  const body =
    '<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:del w:id="1" w:author="Ada">' +
    marker(false) +
    '</w:del><w:ins w:id="2" w:author="Ada">' +
    marker(true).replace('w:val="8"', 'w:val="9"') +
    '</w:ins><w:r><w:t>Task</w:t></w:r></w:p>';
  const value = await imported(body);
  try {
    const { session, revision } = value;
    for (const [read, expected] of [
      [readCommittedView, false],
      [readWorkingView, true],
    ] as const) {
      const result = await read(session, revision);
      if (!result.success) throw new Error(result.error);
      const projection = result.view.projection,
        task = projection.checklists[0]!;
      expect(projection.text).toBe(expected ? '- [x] Task' : '- [ ] Task');
      expect(task.checked).toBe(expected);
      expect(task.native).toEqual(expected ? { start: 1, end: 2 } : { start: 0, end: 1 });
      expect(markdownPointAt(projection, task.contentStart)).toMatchObject({
        paragraphId: task.paragraphId,
        offset: 2,
      });
    }
    const opened = openDocumentForExport(session);
    if (!opened.ok) throw new Error(opened.reason);
    try {
      const all = projectLists(
        session.currentPackage(),
        await opened.session.layoutFor('all-markup'),
        revision
      );
      expect(all.success && all.projection.checklists.length).toBe(1);
      if (all.success) expect(all.projection.checklists[0]!.checked).toBe(true);
    } finally {
      opened.session.dispose();
    }
    expect(session.canUndo()).toBe(false);
  } finally {
    value.editor.destroy();
  }
});

test('explicit protection, locks, bound controls, stale snapshots, and malformed marker ownership refuse before writing', async () => {
  const value = await imported(
    checklistParagraph(false).replace('<w:sdtPr>', '<w:sdtPr><w:lock w:val="contentLocked"/>')
  );
  try {
    const { projection, revision } = value,
      task = projection.checklists[0]!,
      span = { start: task.stateStart, end: task.stateEnd };
    expect(resolvePatchSpan(projection, span, { revision })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(
      resolvePatchSpan(projection, span, {
        revision: { ...revision, revision: revision.revision + 1 },
      })
    ).toMatchObject({ success: false, code: 'stale-revision' });
    expect(
      resolvePatchSpan(projection, span, {
        revision,
        protected: [{ ...span, reason: 'Application read-only region' }],
      })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(value.session.canUndo()).toBe(false);
  } finally {
    value.editor.destroy();
  }
});

test('projection reads preserve native bytes, fingerprints, revision, selection, existing history, and input records', async () => {
  const value = await imported(checklistParagraph(true) + listParagraph('Other'));
  try {
    const { session, editor, layout } = value,
      id = session.paragraphIds()[1]!;
    editor.surface!.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: id, offset: 0, text: 'User ' },
    ]);
    const revision = { ...value.revision, revision: session.packageRevision() };
    const before = writeOoxmlPackage(session.currentPackage()),
      fingerprint = canonicalOoxmlFingerprint(session.part()),
      history = [session.canUndo(), session.canRedo()],
      selection = editor.surface!.state().selection;
    const originalLayout = JSON.stringify(layout);
    const opened = openDocumentForExport(session);
    if (!opened.ok) throw new Error(opened.reason);
    try {
      const fresh = await opened.session.layoutFor('proposed');
      expect(projectLists(session.currentPackage(), fresh, revision).success).toBe(true);
      expect((await readWorkingView(session, revision)).success).toBe(true);
    } finally {
      opened.session.dispose();
    }
    expect(JSON.stringify(layout)).toBe(originalLayout);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(fingerprint);
    expect(session.packageRevision()).toBe(revision.revision);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(editor.surface!.state().selection).toEqual(selection);
  } finally {
    value.editor.destroy();
  }
});
