export const PROPERTY_REVIEW_PROTOTYPE = Symbol.for('sciobot.docx.property-review.prototype.1');
import type { NativeReviewGroup } from './groups';

/** Local canonical evidence. Native IDs stay out of the serialized group binding. */
export interface PropertyGroupScope {
  readonly groupId: string;
  readonly membersKey: string;
  readonly nodeIds: ReadonlySet<string>;
}
export interface PropertyGroupPublication {
  readonly preview: readonly PropertyGroupScope[];
  readonly current: () => readonly PropertyGroupScope[];
}
export function propertyGroupMembersKey(group: NativeReviewGroup): string {
  return JSON.stringify(
    group.native
      .map((member) =>
        JSON.stringify([
          member.partName,
          member.address.id,
          member.address.author,
          member.address.date ?? null,
        ])
      )
      .sort()
  );
}
