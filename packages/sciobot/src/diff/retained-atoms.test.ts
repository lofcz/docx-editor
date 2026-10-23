import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, findNode, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
const context = { documentId: 'retained-atoms', epoch: 'native' };
for (const target of [
  'Keep. $x^2$ [Ada](mention:user-7) <source url="https://example.test" title="Paper" /> Z',
  '$x^2$',
  '**$x^2$**',
  'Keep. **$x^2$** and _[Ada](mention:user-7)_ Z',
  'Keep. [Link](https://example.test)$x^2$',
  'Keep. <scio-inline-equation texExpression="" /> Z',
])
  test(`retained paragraph atom births preserve native ownership, history and DOCX: ${target}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const sourceIds = editor.surface!.session.paragraphIds();
      const paragraphId = sourceIds[0]!;
      const properties = findNode(editor.surface!.session.part(), paragraphId)!;
      if (properties.kind !== 'paragraph') throw Error('Missing paragraph');
      const beforeProperties = properties.children
        .filter((node) => node.kind === 'paragraphProperties')
        .map(canonicalOoxmlFingerprint);
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const makePlan = async () => {
        const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
        const view = await readWorkingView(editor.surface!.session, revision);
        if (!view.success) throw Error(view.error);
        let next = 0;
        const result = planTargetDocument(
          {
            revision,
            pkg: editor.surface!.session.currentPackage(),
            projection: view.view.projection,
          },
          target,
          { allocateBlockId: () => `retained-${++next}` }
        );
        if (!result.success) throw Error(result.error);
        return result.plan;
      };
      const plan = await makePlan();
      expect(plan.operations.some((edit) => edit.kind === 'inline-atoms')).toBe(true);
      expect(plan.operations.some((edit) => edit.kind === 'siblings')).toBe(false);
      expect(decodeEditPlan(serializeEditPlan(plan))).toEqual({ success: true, plan });
      for (const change of ['missing', 'fallback', 'path', 'location', 'offset', 'skeleton']) {
        const copy = JSON.parse(serializeEditPlan(plan));
        const edit = copy.operations.find((edit: { kind: string }) => edit.kind === 'inline-atoms');
        if (change === 'missing') edit.atoms.pop();
        if (change === 'fallback') edit.atoms[0].fallback = 'changed';
        if (change === 'path') edit.atoms[0].path = [999];
        if (change === 'location') edit.atoms[0].location.paragraphId = 'wrong';
        if (change === 'offset') edit.atoms[0].offset += 1;
        if (change === 'skeleton') edit.before += 'changed';
        expect(decodeEditPlan(JSON.stringify(copy)).success).toBe(false);
      }
      const applied = applyDirectPlan(editor, plan, context);
      if (!applied.success) throw Error(applied.error);
      expect(editor.surface!.session.paragraphIds()).toEqual(sourceIds);
      const afterParagraph = findNode(editor.surface!.session.part(), paragraphId)!;
      if (afterParagraph.kind !== 'paragraph') throw Error('Missing retained paragraph');
      expect(
        afterParagraph.children
          .filter((node) => node.kind === 'paragraphProperties')
          .map(canonicalOoxmlFingerprint)
      ).toEqual(beforeProperties);
      expect((await makePlan()).operations).toEqual([]);
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      expect(index.diagnostics).toEqual([]);
      expect(index.byId.has('retained-1')).toBe(true);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      const bytes = new Uint8Array(await editor.save());
      const loaded = readOoxmlPackage(bytes);
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      expect(
        [...buildBlockIndex(loaded.package).byId].map(([id, block]) => [id, block.payload])
      ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
      });
      try {
        const revision = readEditorRevision(reopened, context.documentId, context.epoch)!;
        const view = await readWorkingView(reopened.surface!.session, revision);
        if (!view.success) throw Error(view.error);
        const result = planTargetDocument(
          {
            revision,
            pkg: reopened.surface!.session.currentPackage(),
            projection: view.view.projection,
          },
          target
        );
        expect(result).toMatchObject({ success: true, plan: { operations: [] } });
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

for (const adjacent of [false, true])
  test(`later paragraph edits retain existing atom identities across native text gaps: adjacent=${adjacent}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:r><w:t>Before</w:t></w:r></w:p>'),
    });
    const originalTarget = adjacent
      ? 'Before $x$[Ada](mention:user) <source title="Paper" url="https://example.test" /> after'
      : 'Before $x$ middle [Ada](mention:user) <source title="Paper" url="https://example.test" /> after';
    const target =
      'Start $z$ Before changed $x$PREFIX middle updated [Ada](mention:user) text <source title="Paper" url="https://example.test" /> after End $q$';
    let generation = 0;
    const planFor = async (target: string) => {
      const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const view = await readWorkingView(editor.surface!.session, revision);
      if (!view.success) throw Error(view.error);
      let id = 0;
      const result = planTargetDocument(
        {
          revision,
          pkg: editor.surface!.session.currentPackage(),
          projection: view.view.projection,
        },
        target,
        { allocateBlockId: () => `g${generation}-${++id}` }
      );
      if (!result.success) throw Error(result.error);
      return result.plan;
    };
    try {
      const ids = editor.surface!.session.paragraphIds();
      const initial = applyDirectPlan(editor, await planFor(originalTarget), context);
      if (!initial.success) throw Error(initial.error);
      generation++;
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      const unchanged = [...index.byId].map(([id, block]) => ({
        id,
        payload: block.payload,
        controlId: block.controlId,
        word: canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), block.controlId)!),
      }));
      const planned = await planFor(target);
      expect(planned.operations.filter((edit) => edit.kind === 'inline-atoms')).toHaveLength(2);
      expect(
        planned.operations.some((edit) => edit.kind === 'siblings' || edit.kind === 'payload')
      ).toBe(false);
      expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
      const result = applyDirectPlan(editor, planned, context);
      if (!result.success) throw Error(result.error);
      expect(editor.surface!.session.paragraphIds()).toEqual(ids);
      const current = buildBlockIndex(editor.surface!.session.currentPackage());
      expect(current.diagnostics).toEqual([]);
      expect(current.byId.size).toBe(5);
      for (const item of unchanged) {
        expect(current.byId.get(item.id)?.controlId).toBe(item.controlId);
        expect(current.byId.get(item.id)?.payload).toEqual(item.payload);
        expect(
          canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), item.controlId)!)
        ).toBe(item.word);
      }
      expect((await planFor(target)).operations).toEqual([]);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      expect(
        [...buildBlockIndex(loaded.package).byId].map(([id, block]) => [id, block.payload])
      ).toEqual([...current.byId].map(([id, block]) => [id, block.payload]));
    } finally {
      editor.destroy();
    }
  });

for (const kind of ['heading', 'table'] as const)
  test(`retained ${kind} atom insertion keeps native paragraph and enclosing structure`, async () => {
    const paragraph = '<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>';
    const documentBytes =
      kind === 'heading'
        ? docxFixture(
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Keep.</w:t></w:r></w:p>',
            {
              styles:
                '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>',
            }
          )
        : docxFixture(
            '<w:p><w:r><w:t>Outside</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>H</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr/>' +
              paragraph +
              '</w:tc></w:tr></w:tbl>'
          );
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: documentBytes,
    });
    try {
      const ids = editor.surface!.session.paragraphIds();
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const source = await readWorkingView(editor.surface!.session, revision);
      if (!source.success) throw Error(source.error);
      const target = source.view.projection.text.replace(
        'Keep',
        'Keep $x^2$ [Ada](mention:user) <source url="https://example.test" title="Paper" /> Z'
      );
      let next = 0;
      const result = planTargetDocument(
        {
          revision,
          pkg: editor.surface!.session.currentPackage(),
          projection: source.view.projection,
        },
        target,
        { allocateBlockId: () => `context-${++next}` }
      );
      if (!result.success) throw Error(result.error);
      expect(result.plan.operations.some((edit) => edit.kind === 'inline-atoms')).toBe(true);
      expect(
        result.plan.operations.some((edit) => edit.kind === 'siblings' || edit.kind === 'table')
      ).toBe(false);
      const applied = applyDirectPlan(editor, result.plan, context);
      if (!applied.success) throw Error(applied.error);
      expect(editor.surface!.session.paragraphIds()).toEqual(ids);
      const after = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).diagnostics).toEqual([]);
      const currentRevision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const currentView = await readWorkingView(editor.surface!.session, currentRevision);
      if (!currentView.success) throw Error(currentView.error);
      expect(
        planTargetDocument(
          {
            revision: currentRevision,
            pkg: editor.surface!.session.currentPackage(),
            projection: currentView.view.projection,
          },
          target
        )
      ).toMatchObject({ success: true, plan: { operations: [] } });
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(after);
      const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!loaded.ok) throw Error(loaded.reason);
      expect(
        canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
      ).toBe(after);
      expect(
        [...buildBlockIndex(loaded.package).byId.values()].map((block) => block.payload)
      ).toEqual(
        [...buildBlockIndex(editor.surface!.session.currentPackage()).byId.values()].map(
          (block) => block.payload
        )
      );
    } finally {
      editor.destroy();
    }
  });
