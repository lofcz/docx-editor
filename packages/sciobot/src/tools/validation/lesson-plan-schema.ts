/**
 * Declarative validation rules for the lesson-plan MDX document shape.
 *
 * This is the single source of truth for what a valid complete-preparation
 * document must look like. It is expressed purely as data ({@link
 * MdxValidationSchema}) and applied by the generic {@link validateMdx}
 * engine in validate-mdx.ts. No validation logic lives here — only rules.
 *
 * The tag names mirror `src/components/lesson-plan/types.ts` keys:
 *   PHASE_KEY = "lesson_phase"        → MDX name <phase>
 *   ACTIVITY_KEY = "lesson_activity"  → MDX name <activity>
 *   LESSON_TITLE_KEY / PREPARATION_RECAP_KEY / RECAP_*_KEY → same MDX names.
 * PHASE_TYPES mirrors `src/components/lesson-plan/types.ts#PHASE_TYPES`.
 *
 * Only REQUIRED structure is declared (plus the optional-but-bounded
 * `lesson_visual` nodes the pipeline inserts itself). Unknown / HTML tags,
 * extra attribute names, and expression attribute values are deliberately
 * NOT rules here —
 * the engine tolerates anything it has no rule for, which is the guarantee
 * of zero false positives.
 */

import { type JsxNode, type MdxValidationSchema, structuralChildren } from './schema-engine';

/** Mirrors `src/components/lesson-plan/types.ts#PHASE_TYPES`. */
const PHASE_TYPES = ['evocation', 'understanding', 'reflection'] as const;

/** duration must be a positive whole-minute integer. */
const DURATION_PATTERN = /^[1-9]\d*$/;

/** Mirrors `src/components/lesson-plan/types.ts#LESSON_VISUAL_KINDS`. */
const LESSON_VISUAL_KIND_PATTERN = /^(summary|explainer)$/;

/**
 * Shown when the class declares special needs / an assistant / substitution
 * and the draft has no `<recap_adaptations>` slot. Exported so the validator
 * facade can reuse the exact wording when it tightens the count rule.
 */
export const RECAP_ADAPTATIONS_MISSING_MESSAGE =
  'This class declares special educational needs, a teaching assistant and/or a substitution, so <preparation_recap> must contain one <recap_adaptations> block (plain paragraphs: one row per need labelled with its code, one for the assistant, one for the substitute) placed after <recap_materials> / <recap_curriculum> and before the first <recap_phase>.';

/**
 * The root-level `<lesson_visual>` (summary) is OPTIONAL — `_shared/lesson-
 * visuals.ts` inserts it after the draft — but when present it must sit
 * directly after `</preparation_recap>`, before the first `<phase>`.
 */
function summaryVisualMisplaced(root: JsxNode): boolean {
  const kids = structuralChildren(root);
  const visualIndex = kids.findIndex((k) => k.name === 'lesson_visual');
  if (visualIndex === -1) return false;
  const recapIndex = kids.findIndex((k) => k.name === 'preparation_recap');
  return recapIndex === -1 || visualIndex !== recapIndex + 1;
}

/** A phase explainer `<lesson_visual>` must be the LAST child of its `<phase>`. */
function phaseVisualMisplaced(phase: JsxNode): boolean {
  const kids = structuralChildren(phase);
  const visualIndex = kids.findIndex((k) => k.name === 'lesson_visual');
  return visualIndex !== -1 && visualIndex !== kids.length - 1;
}

export const lessonPlanSchema: MdxValidationSchema = {
  root: {
    checks: [
      {
        code: 'lesson-visual-position',
        message:
          'A root-level <lesson_visual kind="summary" … /> must be placed directly after </preparation_recap> (before the first <phase>).',
        test: summaryVisualMisplaced,
      },
    ],
    children: {
      lesson_title: {
        count: { mode: 'exactly', n: 1 },
        codes: { missing: 'missing-lesson-title' },
        messages: {
          missing:
            'Document must start with a single <lesson_title grade="…">…</lesson_title> element (not a bare # H1).',
        },
      },
      preparation_recap: {
        count: { mode: 'exactly', n: 1 },
        codes: { missing: 'missing-preparation-recap' },
        messages: {
          missing:
            'Document must contain exactly one <preparation_recap> block immediately after <lesson_title>.',
        },
      },
      phase: {
        count: { mode: 'exactly', n: PHASE_TYPES.length },
        codes: { missing: 'phase-count', count: 'phase-count' },
        orderedBy: {
          attribute: 'type',
          sequence: PHASE_TYPES,
          orderCode: 'phase-order',
        },
      },
      // Optional, inserted automatically after the draft (see
      // `_shared/lesson-visuals.ts`): at most ONE lesson summary picture.
      lesson_visual: {
        count: { mode: 'atMost', n: 1 },
        codes: { count: 'duplicate-lesson-visual' },
        messages: {
          count: (found) =>
            `Expected at most one root-level <lesson_visual> (the lesson summary), found ${found}.`,
        },
      },
    },
  },
  containers: {
    // lesson_title carries the grade badge. Whether it is spelled as a
    // string, an expression, a boolean, or a differently-typed attr, the
    // renderer degrades gracefully — so we deliberately impose no attribute
    // rule here and only require the element itself to be present.
    lesson_title: {},
    preparation_recap: {
      children: {
        recap_goals: {
          count: { mode: 'exactly', n: 1 },
          codes: { missing: 'missing-recap-goals' },
          messages: {
            missing: '<preparation_recap> must contain a <recap_goals> slot.',
          },
        },
        recap_phase: {
          count: { mode: 'exactly', n: PHASE_TYPES.length },
          codes: { missing: 'recap-phase-count', count: 'recap-phase-count' },
          orderedBy: {
            attribute: 'type',
            sequence: PHASE_TYPES,
            orderCode: 'recap-phase-order',
          },
        },
        // Optional by default; `validateLessonPlanMarkdown` tightens this to
        // exactly one when the class declares SEN / assistant / substitution.
        recap_adaptations: {
          count: { mode: 'atMost', n: 1 },
          codes: { missing: 'missing-recap-adaptations', count: 'duplicate-recap-adaptations' },
          messages: {
            missing: RECAP_ADAPTATIONS_MISSING_MESSAGE,
            count: (found) =>
              `Expected at most one <recap_adaptations> block inside <preparation_recap>, found ${found}. Merge the rows into a single block.`,
          },
        },
      },
    },
    recap_goals: {},
    recap_adaptations: {},
    recap_phase: {
      attributes: { type: { required: true } },
    },
    phase: {
      attributes: { type: { required: true } },
      checks: [
        {
          code: 'empty-phase',
          message:
            'A phase must contain an activity unless it explicitly has duration="0" (a continuation with no time allocated to this phase).',
          test: (phase) =>
            structuralChildren(phase).every((child) => child.name !== 'activity') &&
            phase.attributes?.find((attribute) => attribute.name === 'duration')?.value !== '0',
        },
        {
          code: 'lesson-visual-position',
          message:
            'A <lesson_visual> inside a <phase> must be its LAST child, after every <activity>.',
          test: phaseVisualMisplaced,
        },
      ],
      children: {
        activity: {
          count: { mode: 'atLeast', n: 0 },
          codes: { missing: 'empty-phase' },
          messages: {
            missing: '<phase> must contain at least one <activity>.',
          },
        },
        // Optional phase explainer picture, inserted automatically.
        lesson_visual: {
          count: { mode: 'atMost', n: 1 },
          codes: { count: 'duplicate-lesson-visual' },
          messages: {
            count: (found) => `Expected at most one <lesson_visual> per <phase>, found ${found}.`,
          },
        },
      },
    },
    lesson_visual: {
      attributes: {
        kind: {
          required: true,
          pattern: LESSON_VISUAL_KIND_PATTERN,
          missingCode: 'lesson-visual-missing-kind',
          missingMessage: `<lesson_visual> is missing kind="summary|explainer".`,
          invalidCode: 'lesson-visual-invalid-kind',
          invalidMessage: (value) =>
            `<lesson_visual> has invalid kind="${value}". Use "summary" or "explainer".`,
        },
        src: {
          required: true,
          nonEmpty: true,
          missingCode: 'lesson-visual-missing-src',
          missingMessage: `<lesson_visual> is missing a non-empty src="…" image URL.`,
          invalidCode: 'lesson-visual-missing-src',
          invalidMessage: `<lesson_visual> has an empty src="" — it needs an image URL.`,
        },
      },
    },
    activity: {
      attributes: {
        name: {
          required: true,
          nonEmpty: true,
          missingCode: 'activity-missing-name',
          missingMessage: `<activity> is missing a non-empty name="…" attribute.`,
        },
        duration: {
          required: true,
          nonEmpty: true,
          pattern: DURATION_PATTERN,
          missingCode: 'activity-missing-duration',
          missingMessage: `<activity> is missing duration="…". duration must be a positive whole-minute integer (e.g. duration="10").`,
          invalidCode: 'activity-invalid-duration',
          invalidMessage: (value) =>
            `<activity> has invalid duration="${value}". Use a positive whole-minute integer (e.g. duration="10").`,
        },
      },
      checks: [
        {
          code: 'activity-empty',
          message:
            '<activity> has no body. An activity must contain a bold opener paragraph followed by teachable content.',
          test: (node) => node.type === 'mdxJsxFlowElement' && (node.children ?? []).length === 0,
        },
      ],
    },
  },
};
