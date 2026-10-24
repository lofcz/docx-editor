import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { writeOoxmlPackage } from '@docx-editor.dev/core/store';
import {
  atoms,
  cases,
  context,
  decodePort,
  png,
  preparedCellDocument,
} from '../../test/support/tracked-cell-atoms';
import { readEditorRevision } from '../projection/revision-cache';
import { recomposeOwnedProposals } from './recompose';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readPropertyReviewStore } from './property-record-store';
import { decodeBlockPresenceRecord, type BlockPresenceRecord } from './block-presence-schema';
import { bindBlockPresenceRecord } from './block-presence-plan';

for (const item of [
  { name: 'cell', body: cases[0]!.body, target: cases[0]!.target, profile: 5 as const },
  {
    name: 'whole table',
    body: '<w:p><w:r><w:t>After</w:t></w:r></w:p>',
    target: `| A | B |\n| - | - |\n| ${atoms} | Tail |\n\nAfter`,
    profile: 6 as const,
  },
])
  test(`${item.name} birth cannot be rebound as a removal proposal`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: preparedCellDocument(item.body),
      imageDecodePort: decodePort,
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    try {
      const staged = await recomposeOwnedProposals(
        editor,
        {
          revision: readEditorRevision(editor, context.documentId, context.epoch)!,
          targetMarkdown: item.target,
          ownership: [],
        },
        context,
        {
          payloads: { presence: { proposalId: 'birth-direction' } },
          images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
        }
      );
      if (!staged.success) throw Error(staged.error);
      const pkg = editor.surface!.session.currentPackage(),
        bytes = writeOoxmlPackage(pkg);
      const store = readPropertyReviewStore(pkg);
      if (store.status !== 'ready') throw Error('Missing proposal store.');
      const decoded = decodeBlockPresenceRecord(store.records.get('birth-direction')!.raw);
      if (decoded.status !== 'supported') throw Error('Missing birth proposal.');
      expect(decoded.value.schemaVersion).toBe(item.profile);
      expect(bindBlockPresenceRecord(pkg, decoded.value).success).toBe(true);
      const forged: { blocks: { before: { present: boolean }; after: { present: boolean } }[] } =
        JSON.parse(JSON.stringify(decoded.value));
      for (const block of forged.blocks) {
        block.before.present = true;
        block.after.present = false;
      }
      expect(bindBlockPresenceRecord(pkg, forged as unknown as BlockPresenceRecord).success).toBe(
        false
      );
      expect(writeOoxmlPackage(pkg)).toEqual(bytes);
    } finally {
      editor.destroy();
    }
  }, 30000);
