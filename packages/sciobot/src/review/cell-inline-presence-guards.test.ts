import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { cases, context, decodePort, png } from '../../test/support/tracked-cell-atoms';
import { readEditorRevision } from '../projection/revision-cache';
import { recomposeOwnedProposals } from './recompose';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readPropertyReviewStore } from './property-record-store';
import { decodeBlockPresenceRecord, type BlockPresenceRecord } from './block-presence-schema';
import { bindBlockPresenceRecord } from './block-presence-plan';
import { readInlinePresenceExtension } from './inline-presence-native';

interface MutableCellProof {
  schemaVersion: number;
  extension: {
    nativeInline: {
      schemaVersion: number;
      owners: [
        {
          revision: { id: string; author: string; date?: string };
          siblings?: true;
          cell: {
            kind: string;
            revision: { id: string; author: string; date?: string };
            tableId?: string;
          };
          gap: [
            { id: string; author: string; date?: string },
            ...{ id: string; author: string; date?: string }[],
          ];
        },
        ...unknown[],
      ];
    };
  };
}

test('cell presence binds persistent topology addresses and refuses changed or downgraded proofs without writes', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(cases[0]!.body),
    imageDecodePort: decodePort,
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  try {
    const staged = await recomposeOwnedProposals(
      editor,
      {
        revision: readEditorRevision(editor, context.documentId, context.epoch)!,
        targetMarkdown: cases[0]!.target,
        ownership: [],
      },
      context,
      {
        payloads: { presence: { proposalId: 'cell-guard' } },
        images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
      }
    );
    if (!staged.success) throw Error(staged.error);
    const pkg = editor.surface!.session.currentPackage(),
      store = readPropertyReviewStore(pkg);
    if (store.status !== 'ready') throw Error('Missing presence store.');
    const decoded = decodeBlockPresenceRecord(store.records.get('cell-guard')!.raw);
    if (decoded.status !== 'supported') throw Error('Missing supported presence.');
    const record = decoded.value,
      bytes = writeOoxmlPackage(pkg),
      owners = readInlinePresenceExtension(record)!;
    expect(record.schemaVersion).toBe(5);
    expect(owners.size).toBe(4);
    for (const owner of owners.values()) {
      expect(owner.cell?.kind).toBe('row');
      expect(owner.cell?.revision.author).toBe('AI');
      expect(owner.gap!.some((address) => address.id === owner.cell!.revision.id)).toBe(true);
      expect(Object.keys(owner.cell!).sort()).toEqual(['kind', 'revision']);
    }
    expect(bindBlockPresenceRecord(pkg, record).success).toBe(true);
    const change = (update: (value: MutableCellProof) => void): BlockPresenceRecord => {
      const value: MutableCellProof = JSON.parse(JSON.stringify(record));
      update(value);
      return value as unknown as BlockPresenceRecord;
    };
    const invalid = [
      change((value) => {
        value.schemaVersion = 6;
        value.extension.nativeInline.schemaVersion = 4;
        for (const owner of value.extension.nativeInline.owners) {
          const forged = owner as { cell?: unknown; siblings?: true };
          delete forged.cell;
          forged.siblings = true;
        }
      }),
      change(
        (value) => (value.extension.nativeInline.owners[0].cell.revision.id = 'missing-topology')
      ),
      change((value) => (value.extension.nativeInline.owners[0].cell.revision.author = 'Other')),
      change((value) => (value.extension.nativeInline.owners[0].cell.kind = 'cell')),
      change((value) => (value.extension.nativeInline.owners[0].revision.id = 'missing-paragraph')),
      change((value) => value.extension.nativeInline.owners[0].gap.pop()),
      change((value) =>
        value.extension.nativeInline.owners[0].gap.push(
          value.extension.nativeInline.owners[0].gap[0]
        )
      ),
      change((value) => (value.extension.nativeInline.owners[0].cell.tableId = 'transient-id')),
      change((value) => {
        value.schemaVersion = 4;
        value.extension.nativeInline.schemaVersion = 2;
      }),
      change((value) => {
        value.schemaVersion = 3;
        value.extension.nativeInline.schemaVersion = 1;
      }),
    ];
    for (const [index, value] of invalid.entries()) {
      expect(bindBlockPresenceRecord(pkg, value).success, `invalid cell proof ${index}`).toBe(
        false
      );
      expect(writeOoxmlPackage(pkg)).toEqual(bytes);
    }
  } finally {
    editor.destroy();
  }
});
