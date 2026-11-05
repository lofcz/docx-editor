import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import {
  addsTaskCompletionStyles,
  TASK_COMPLETION_PREPARATION_ERROR,
} from '../native/task-completion-style';
import { refuseWrite } from '../diff/fragment-seeds';

export function assertTaskCompletionPrepared(
  editor: DocxEditorInstance,
  before: OoxmlPackage,
  after: OoxmlPackage
): void {
  if (
    editor.snapshot().collaborationStatus !== 'inactive' &&
    addsTaskCompletionStyles(before, after)
  )
    refuseWrite(TASK_COMPLETION_PREPARATION_ERROR, 'unsupported-operation');
}
