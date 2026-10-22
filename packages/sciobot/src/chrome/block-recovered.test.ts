// The shared regression uses the EigenPal Pro License transport.
import { expect, test } from 'bun:test';
import {
  directParagraphProperties,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { blockContext, decideBlocks, manualBlockPlan } from '../../test/support/block-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { trackSiblingPlan } from '../review/track-blocks';
import { withBlockTypeStyles } from './block-styles';
import { setBlockType } from './block-type';

test('recovered source uses its selected heading for Enter instead of the removed heading mark', async () => {
  const loaded = readOoxmlPackage(
    docxFixture('<w:p><w:r><w:t>Opening</w:t></w:r></w:p><w:p><w:r><w:t>Closing</w:t></w:r></w:p>')
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const pair = await createTwoEditorFixture({
    document: writeOoxmlPackage(withBlockTypeStyles(loaded.package)!),
    clientIds: [41, 73],
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
  try {
    const [alice, bob] = pair.editors,
      a = alice!.surface!,
      b = bob!.surface!;
    const tracked = trackSiblingPlan(alice!, manualBlockPlan(alice!, 1, 0, '# New'), blockContext);
    if (!tracked.success) throw new Error(tracked.error);
    await pair.settle();
    const id = b.session.paragraphIds()[1]!;
    pair.pause();
    decideBlocks(alice!, tracked.proposals, 'rejectRevision');
    expect(
      b.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' }],
        { kind: 'body' },
        undefined,
        false,
        true
      ).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    expect(paragraphTextOf(a.session.part(), id)).toBe('Peer');
    a.setSelection({
      anchor: { paragraphId: id, offset: 2 },
      head: { paragraphId: id, offset: 2 },
    });
    expect(
      setBlockType(alice!, 'h2', { documentId: blockContext.documentId, epoch: blockContext.epoch })
        .success
    ).toBe(true);
    a.splitParagraph();
    await pair.settle();
    for (const surface of [a, b]) {
      const ids = surface.session.paragraphIds();
      expect(ids.map((value) => paragraphTextOf(surface.session.part(), value))).toEqual([
        'Opening',
        'Pe',
        'er',
        'Closing',
      ]);
      expect(
        ids
          .slice(1, 3)
          .map(
            (value) =>
              directParagraphProperties(surface.session.part(), value).find(
                (p) => p.localName === 'pStyle'
              )?.attributes?.val
          )
      ).toEqual(['Heading2', 'Heading2']);
    }
    const saved = await pair.savedDigests();
    expect(saved[0]).toEqual(saved[1]);
  } finally {
    pair.destroy();
  }
});
