import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  directParagraphProperties,
  bodyStoryRoot,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  paraIdOf,
  storyParagraphs,
  readOoxmlPart,
  withPart,
  type OoxmlPackage,
  type TreeDocOp,
  type TreeTransactionContext,
} from '@docx-editor.dev/core/store';
import { nativeNumberingInfrastructureReady } from '../native/numbering-infrastructure';
import { applyDirectListAssignments } from '../ai/direct-lists';
import { listInstanceLevel, planListMembership } from '../diff/lists';
import { readEditorRevision } from '../projection/revision-cache';
import { stageTrackedChecklist } from '../review/tracked-stage-checklists';
import { captureReviewGroupHistory } from '../review/group-history';
import { withReviewGroupRevisionReservations } from '../review/group-allocation';
import type { ListTarget } from '../diff/list-types';
import type { TrackedStageWriter } from '../review/tracked-stage-writer';
import type { BlockCommandContext } from './block-type';
import { checklistSplitBudget, persistListGestureGroups } from './list-split-tracking';
const fail = (error: string) => ({
  success: false as const,
  code: 'unsupported-operation' as const,
  error,
});

/**
 * Insert a list block at the native replacement landing, preserving unselected Word content.
 * paragraphId addresses this revision; paragraphIdentity follows later collaboration remapping.
 */
export function insertManualList(
  editor: DocxEditorInstance,
  style: 'disc' | 'decimal' | 'todo',
  context: BlockCommandContext
) {
  const s = editor.surface;
  if (!s) return fail('The native editor is not attached.');
  s.flushPendingInput();
  if (!['disc', 'decimal', 'todo'].includes(style)) return fail('Choose an insertion list style.');
  const original = s.session.currentPackage(),
    source = s.session.partFor(s.storyScope()) ?? s.session.part();
  if (source.name !== original.mainDocumentPart)
    return fail('List insertion requires the body story.');
  if (editor.getTableCellSelection()) return fail('List insertion requires a text selection.');
  if (
    style !== 'todo' &&
    editor.snapshot().collaborationStatus !== 'inactive' &&
    !nativeNumberingInfrastructureReady(original)
  )
    return fail('This document is not ready for shared list insertion.');
  const revision = readEditorRevision(editor, context.documentId, context.epoch);
  if (!revision) return fail('The native document is not ready.');
  const selection = s.state().selection,
    order = s.session.paragraphIds();
  const a = order.indexOf(selection.anchor.paragraphId),
    h = order.indexOf(selection.head.paragraphId);
  const first =
    a < h || (a === h && selection.anchor.offset <= selection.head.offset)
      ? selection.anchor
      : selection.head;
  const removesHead =
    first.offset === 0 && (a !== h || selection.anchor.offset !== selection.head.offset);
  const plan = s.planRangeDeletion(selection),
    suggesting = s.editingMode() === 'suggest';
  const parsed = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
    { name: source.name, contentType: source.contentType }
  );
  if (!parsed.ok) return fail('The native insertion seed is invalid.');
  const body = bodyStoryRoot(parsed.part),
    seed = body && body.kind !== 'textValue' ? body.children[0] : null;
  if (seed?.kind !== 'paragraph') return fail('The native insertion seed has no paragraph.');
  const attribution = {
    author: s.author() ?? '',
    date: `${new Date().toISOString().slice(0, 19)}Z`,
  };
  let paragraphIdentity: string | null = null;
  let inserted: string | null = null,
    caretOffset = style === 'todo' ? 1 : 0;
  const target: ListTarget =
    style === 'todo'
      ? { kind: 'checklist', checked: false, level: 0 }
      : {
          kind: style === 'disc' ? 'bullet' : 'ordered',
          style,
          level: 0,
          instance: { kind: 'new', key: `insert:${revision.revision}` },
        };
  const complete = (
    pkg: OoxmlPackage,
    measure?: (before: OoxmlPackage, after: OoxmlPackage) => void,
    stage?: TreeTransactionContext
  ) => {
    let working = pkg;
    const writer: TrackedStageWriter = {
      currentPackage: () => working,
      apply(partName, op) {
        const prior = working,
          part = working.parts.get(partName)!;
        if (stage) {
          if (!stage.applyTo(partName, op)) throw new Error(`List insertion refused: ${op.op}.`);
          working = stage.currentPackage();
        } else {
          const result = applyTreeOp(part, op);
          if (!result.ok) throw new Error(`List insertion refused: ${result.reason}.`);
          working = withPart(working, result.part);
        }
        measure?.(prior, working);
      },
      edit(change) {
        const prior = working;
        if (stage) {
          if (!stage.applyPackage(change)) throw new Error('The list package edit refused.');
          working = stage.currentPackage();
        } else working = change(working);
        measure?.(prior, working);
      },
    };
    const part = working.parts.get(source.name)!,
      point = plan.replaceAt,
      paragraph = findNode(part, point.paragraphId),
      parent = parentNodeOf(part, point.paragraphId),
      text = paragraphTextOf(part, point.paragraphId);
    if (paragraph?.kind !== 'paragraph' || !parent || text === null || point.offset > text.length)
      throw new Error('The native insertion landing disappeared.');
    const inspected = planListMembership({ pkg: working, partName: source.name, revision }, [
      { paragraphId: paragraph.id, target: { kind: 'none' } },
    ]);
    if (!inspected.success) throw new Error(inspected.error);
    const membership = inspected.plan.operations.find((edit) => edit.kind === 'list')?.before;
    const numbering = membership?.numbering;
    const sourceDefinition = numbering
      ? listInstanceLevel(working, numbering.numId, numbering.level)?.level
      : undefined;
    const rootDefinition =
      numbering?.level === 0
        ? sourceDefinition
        : numbering
          ? listInstanceLevel(working, numbering.numId, 0)?.level
          : undefined;
    const compatibleRoot =
      !membership?.checkbox &&
      style !== 'todo' &&
      (style === 'decimal'
        ? rootDefinition?.numFmt === 'decimal'
        : rootDefinition?.numFmt === 'bullet' && rootDefinition.lvlText === '•');
    const sourceFormat = sourceDefinition?.numFmt;
    // The captured insertion dispatcher classifies every non-decimal marker as a bullet block.
    const sameStyle =
      style === 'todo'
        ? !!membership?.checkbox
        : !membership?.checkbox &&
          !!sourceFormat &&
          sourceFormat !== 'none' &&
          (style === 'decimal' ? sourceFormat === 'decimal' : sourceFormat !== 'decimal');
    const empty =
      !text.length ||
      (!!membership?.checkbox && text.length === 1) ||
      (suggesting && removesHead && point.offset === text.length);
    const appendEmpty = empty && sameStyle;
    const membershipTarget: ListTarget =
      target.kind !== 'checklist' && compatibleRoot
        ? { kind: target.kind, level: 0, instance: { kind: 'existing', numId: numbering!.numId } }
        : target;
    // A wholly replaced head becomes the inserted block. Its old content remains struck until review.
    const reuse = !appendEmpty && (empty || (suggesting && removesHead && point.offset > 0));
    if (reuse) {
      if (suggesting && point.offset > 0 && point.offset < text.length) {
        writer.apply(source.name, {
          op: 'splitParagraph',
          paragraphId: paragraph.id,
          offset: point.offset,
        });
        writer.apply(source.name, {
          op: 'setParagraphMarkRevision',
          paragraphId: paragraph.id,
          kind: 'ins',
          revision: attribution,
        });
      }
      inserted = paragraph.id;
      const properties = directParagraphProperties(working.parts.get(source.name)!, paragraph.id);
      if (properties.some((property) => property.localName === 'pStyle'))
        writer.apply(source.name, {
          op: 'setParagraphProperties',
          paragraphId: paragraph.id,
          properties: properties.filter((property) => property.localName !== 'pStyle'),
          ...(suggesting ? { revision: attribution } : {}),
        });
    } else {
      let beforeId: string | null;
      if (point.offset === 0 && !appendEmpty) beforeId = paragraph.id;
      else {
        if (point.offset < text.length && !appendEmpty) {
          writer.apply(source.name, {
            op: 'splitParagraph',
            paragraphId: paragraph.id,
            offset: point.offset,
          });
          if (suggesting)
            writer.apply(source.name, {
              op: 'setParagraphMarkRevision',
              paragraphId: paragraph.id,
              kind: 'ins',
              revision: attribution,
            });
        }
        const currentParent = findNode(working.parts.get(source.name)!, parent.id);
        if (!currentParent || currentParent.kind === 'textValue')
          throw new Error('The insertion parent disappeared.');
        const at = currentParent.children.findIndex((node) => node.id === paragraph.id);
        beforeId = currentParent.children[at + 1]?.id ?? null;
      }
      const currentParent = findNode(working.parts.get(source.name)!, parent.id);
      if (!currentParent || currentParent.kind === 'textValue')
        throw new Error('The insertion parent disappeared.');
      const priorIds = new Set(currentParent.children.map((node) => node.id));
      const operation: TreeDocOp = {
        op: 'insertParagraphs',
        parentId: parent.id,
        beforeId,
        paragraphs: [seed],
        ...(suggesting ? { revision: attribution } : {}),
      };
      writer.apply(source.name, operation);
      const afterParent = findNode(working.parts.get(source.name)!, parent.id);
      if (!afterParent || afterParent.kind === 'textValue')
        throw new Error('The insertion parent disappeared.');
      inserted =
        afterParent.children.find((node) => node.kind === 'paragraph' && !priorIds.has(node.id))
          ?.id ?? null;
      if (!inserted) throw new Error('The list paragraph was not allocated.');
    }
    const list = planListMembership({ pkg: working, partName: source.name, revision }, [
      { paragraphId: inserted!, target: membershipTarget },
    ]);
    if (!list.success) throw new Error(list.error);
    for (const edit of list.plan.operations) {
      if (edit.kind !== 'list') throw new Error('The list membership plan is invalid.');
      if (suggesting && (style === 'todo' || edit.before.checkbox))
        stageTrackedChecklist(writer, list.plan, edit, attribution);
      else
        writer.edit((value) =>
          applyDirectListAssignments(
            value,
            source.name,
            [{ paragraphId: edit.paragraphId, before: edit.before, target: edit.after }],
            undefined,
            suggesting ? attribution : undefined
          )
        );
    }
    caretOffset =
      suggesting && reuse
        ? paragraphTextOf(working.parts.get(source.name)!, inserted!)!.length
        : style === 'todo'
          ? 1
          : 0;
    if (suggesting)
      writer.edit((value) =>
        persistListGestureGroups(
          original,
          value,
          source.name,
          editor.snapshot().collaborationStatus === 'inactive'
        )
      );
    const allocated = inserted && findNode(working.parts.get(source.name)!, inserted);
    paragraphIdentity = allocated ? paraIdOf(allocated) : null;
    if (!paragraphIdentity) throw new Error('The inserted paragraph has no Word identity.');
    return working;
  };
  const resolveInserted = () => {
    const part = s.session.part(),
      body = bodyStoryRoot(part);
    return paragraphIdentity && body
      ? (storyParagraphs(body).find((node) => paraIdOf(node) === paragraphIdentity)?.id ?? null)
      : null;
  };
  let committed = false;
  try {
    const result = withReviewGroupRevisionReservations(original, () => {
      const budget = checklistSplitBudget(original, source.name, plan.ops, complete);
      return s.commitReviewOps(
        () => {
          const execute = () =>
            s.applyAutomationOps(
              () => plan.ops,
              { kind: 'body' },
              undefined,
              suggesting,
              true,
              'after',
              (ctx) => {
                complete(ctx.currentPackage(), undefined, ctx);
              }
            );
          const applied = suggesting
            ? captureReviewGroupHistory(
                s.collaborationSession(),
                execute,
                budget.flags,
                budget.fields
              )
            : execute();
          committed = applied.committed && !applied.rejected;
          return applied;
        },
        'revision-resolve',
        () => {
          if (!committed || !inserted) return null;
          const current = resolveInserted();
          if (!current) return null;
          const point = { paragraphId: current, offset: caretOffset };
          return { anchor: point, head: point };
        }
      );
    });
    if (result !== null) return fail(typeof result === 'string' ? result : result.error);
    const current = resolveInserted();
    return committed && current && paragraphIdentity
      ? { success: true as const, paragraphId: current, paragraphIdentity }
      : fail('The list insertion did not commit.');
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
}
