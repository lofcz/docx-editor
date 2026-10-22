import { reviewGroupRevisionReservations } from './group-allocation';
import type { EditorModule } from '@docx-editor.dev/core/editor';
import {
  collectReviewItems,
  findNode,
  revisionItemsOf,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';

/** Application-owned review contribution over Apache core reads. It contains no Pro runtime. */
export const SCIOBOT_NATIVE_REVIEW_MODULE: EditorModule = Object.freeze({
  id: 'sciobot-native-review',
  review: Object.freeze({
    reservedRevisionIds: reviewGroupRevisionReservations,
    displayModes: Object.freeze(['all-markup', 'simple-markup', 'proposed', 'original'] as const),
    collectReviewItems,
    revisionItemsOfParagraph: (part: OoxmlPart, id: string) => {
      const root = findNode(part, id);
      return root?.kind === 'paragraph' ? revisionItemsOf({ ...part, root }) : [];
    },
  }),
});
