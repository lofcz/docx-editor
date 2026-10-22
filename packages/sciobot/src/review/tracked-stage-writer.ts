import type { OoxmlPackage, TreeDocOp } from '@docx-editor.dev/core/store';

/** Every transition passes the replacement coordinator's admission and rollback boundary. */
export interface TrackedStageWriter {
  currentPackage(): OoxmlPackage;
  apply(partName: string, op: TreeDocOp): void;
  edit(change: (pkg: OoxmlPackage) => OoxmlPackage): void;
}

export interface TrackedStageAttribution {
  readonly author: string;
  readonly date: string;
}
