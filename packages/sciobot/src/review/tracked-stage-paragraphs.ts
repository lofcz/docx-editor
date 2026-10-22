import { ensureHyperlinkRelationship } from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import {
  inlineCodeCharacterStyles,
  inlineCodeRunChanges,
  withInlineCodeStyle,
} from '../native/inline-code';
import type { newParagraphSeeds } from './new-paragraphs';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageWriter } from './tracked-stage-writer';

type Resources = Pick<
  Extract<ReturnType<typeof newParagraphSeeds>, { success: true }>,
  'links' | 'codes'
>;
/** Add native resources only after this stage allocates its new owned paragraphs. */
export function stageParagraphResources(
  writer: TrackedStageWriter,
  partName: string,
  paragraphIds: readonly string[],
  seeds: Resources
): void {
  for (const link of seeds.links) {
    let target: { relationshipId: string } | { anchor: string };
    if (link.target.kind === 'bookmark') target = { anchor: link.target.anchor };
    else {
      const url = link.target.url;
      let relationshipId = '';
      writer.edit((pkg) => {
        const resource = ensureHyperlinkRelationship(pkg, url, partName);
        if (!resource)
          refuseWrite('A new native hyperlink target was refused.', 'unsupported-operation');
        relationshipId = resource.relationshipId;
        return resource.pkg;
      });
      target = { relationshipId };
    }
    writer.apply(partName, {
      op: 'insertHyperlink',
      paragraphId: paragraphIds[link.paragraphIndex]!,
      start: link.start,
      end: link.end,
      ...target,
      ...(link.title !== null ? { tooltip: link.title } : {}),
    });
  }
  if (seeds.codes.length)
    writer.edit((pkg) => {
      const styled = withInlineCodeStyle(pkg);
      if (!styled)
        refuseWrite('The document cannot gain the inline code style.', 'unsupported-operation');
      return styled;
    });
  for (const code of seeds.codes) {
    const paragraphId = paragraphIds[code.paragraphIndex]!;
    const changes = inlineCodeRunChanges(
      currentStagePart(writer, partName),
      {
        paragraphId,
        start: code.start,
        end: code.end,
        code: true,
      },
      { codeStyles: inlineCodeCharacterStyles(writer.currentPackage()) }
    );
    if (!changes.success) refuseWrite(changes.error, changes.code);
    for (const change of changes.changes) {
      writer.edit((pkg) => {
        const styled = withInlineCodeStyle(pkg, change.after === null ? [] : [change.after]);
        if (!styled)
          refuseWrite('The document cannot gain the inline code style.', 'unsupported-operation');
        return styled;
      });
      writer.apply(partName, {
        op: 'setRunProperties',
        paragraphId,
        start: change.start,
        end: change.end,
        properties: change.properties,
        characterStyle: change.after,
      });
    }
  }
}
