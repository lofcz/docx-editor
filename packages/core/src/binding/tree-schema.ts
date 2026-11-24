import type { Schema } from 'prosemirror-model';
import { getTreeSchema } from './tree-schema-definition.ts';

export { runPropsOf, type ParagraphAttrs } from './tree-schema-definition.ts';

/** The canonical projection schema shared by every binding session. */
export const treeSchema: Schema<
  'doc' | 'paragraph' | 'text' | 'tab' | 'hardBreak' | 'pageBreak' | 'unknownInline',
  'runProps'
> = /* @__PURE__ */ getTreeSchema();
