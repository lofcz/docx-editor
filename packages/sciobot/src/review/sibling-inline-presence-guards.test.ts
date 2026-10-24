import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  collectRevisionSites,
  bodyStoryRoot,
  applyTreeOp,
  readOoxmlPart,
  parentNodeOf,
  replaceNode,
  withPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  atoms,
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
import { readInlinePresenceExtension } from './inline-presence-native';

interface MutableProof {
  schemaVersion: number;
  extension: {
    nativeInline: {
      schemaVersion: number;
      owners: {
        revision: { id: string; author: string; date?: string };
        siblings?: unknown;
        gap: { id: string; author: string; date?: string }[];
      }[];
    };
  };
}
test('whole-table sibling presence refuses incomplete, forged, and downgraded receipts without writes', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: preparedCellDocument(
      '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>Old</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>'
    ),
    imageDecodePort: decodePort,
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  try {
    const staged = await recomposeOwnedProposals(
      editor,
      {
        revision: readEditorRevision(editor, context.documentId, context.epoch)!,
        targetMarkdown: `Before\n\n| A | B |\n| - | - |\n| ${atoms} | Tail |\n\nAfter`,
        ownership: [],
      },
      context,
      {
        payloads: { presence: { proposalId: 'sibling-guards' } },
        images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
      }
    );
    if (!staged.success) throw Error(staged.error);
    const pkg = editor.surface!.session.currentPackage(),
      store = readPropertyReviewStore(pkg);
    if (store.status !== 'ready') throw Error('Missing proposal store.');
    const decoded = decodeBlockPresenceRecord(store.records.get('sibling-guards')!.raw);
    if (decoded.status !== 'supported') throw Error('Missing whole-table proposal.');
    const record = decoded.value,
      bytes = writeOoxmlPackage(pkg);
    expect(record.schemaVersion).toBe(6);
    const owners = readInlinePresenceExtension(record)!;
    expect(owners.size).toBe(4);
    for (const owner of owners.values()) {
      expect(owner.siblings).toBe(true);
      expect(owner.gap).toHaveLength(2);
      expect(Object.keys(owner).sort()).toEqual([
        'blockId',
        'gap',
        'kind',
        'partName',
        'revision',
        'siblings',
      ]);
    }
    expect(bindBlockPresenceRecord(pkg, record).success).toBe(true);
    const change = (update: (value: MutableProof) => void): BlockPresenceRecord => {
      const value: MutableProof = JSON.parse(JSON.stringify(record));
      update(value);
      return value as unknown as BlockPresenceRecord;
    };
    const invalid = [
      change((v) => {
        v.extension.nativeInline.owners[0]!.siblings = false;
      }),
      change((v) => {
        v.extension.nativeInline.owners[0]!.siblings = { tableId: 'transient' };
      }),
      change((v) => {
        v.extension.nativeInline.owners[0]!.gap.pop();
      }),
      change((v) => {
        const owner = v.extension.nativeInline.owners[0]!;
        owner.gap.push(owner.gap[0]!);
      }),
      change((v) => {
        v.extension.nativeInline.owners[0]!.gap[0]!.author = 'Other';
      }),
      change((v) => {
        v.extension.nativeInline.owners[0]!.revision.id = 'missing';
      }),
      ...[3, 4, 5].map((version) =>
        change((v) => {
          v.schemaVersion = version;
          v.extension.nativeInline.schemaVersion = version - 2;
          v.extension.nativeInline.owners.forEach((o) => {
            delete o.siblings;
          });
        })
      ),
    ];
    for (const [index, candidate] of invalid.entries()) {
      expect(
        bindBlockPresenceRecord(pkg, candidate).success,
        `invalid sibling proof ${index}`
      ).toBe(false);
      expect(writeOoxmlPackage(pkg)).toEqual(bytes);
    }
    const part = pkg.parts.get(pkg.mainDocumentPart)!;
    const body = bodyStoryRoot(part);
    if (!body || body.kind === 'textValue') throw Error('Missing body.');
    const table = body.children.find((node) => node.kind === 'table');
    const seed = readOoxmlPart(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Unowned</w:t></w:r></w:p>',
      { name: '/unowned.xml', contentType: part.contentType }
    );
    if (!table || !seed.ok || seed.part.root.kind !== 'paragraph')
      throw Error('Missing interleaved root fixture.');
    const interleaved = applyTreeOp(part, {
      op: 'insertParagraphs',
      parentId: body.id,
      beforeId: table.id,
      paragraphs: [seed.part.root],
    });
    if (!interleaved.ok) throw Error(interleaved.reason);
    const unrelated = withPart(pkg, interleaved.part),
      unrelatedBytes = writeOoxmlPackage(unrelated);
    expect(bindBlockPresenceRecord(unrelated, record).success).toBe(false);
    expect(writeOoxmlPackage(unrelated)).toEqual(unrelatedBytes);

    const rowBirth = collectRevisionSites(part).find(
      (s) => s.node.localName === 'ins' && parentNodeOf(part, s.node.id)?.localName === 'trPr'
    );
    if (!rowBirth) throw Error('Missing row birth.');
    const parent = parentNodeOf(part, rowBirth.node.id)!;
    if (parent.kind !== 'generic') throw Error('Invalid row property holder.');
    const removed = replaceNode(part, parent.id, {
      ...parent,
      children: parent.children.filter((n) => n.id !== rowBirth.node.id),
    });
    if (!removed.ok) throw Error('Cannot construct unowned row.');
    const forged = withPart(pkg, removed.part),
      snapshot = writeOoxmlPackage(forged);
    expect(bindBlockPresenceRecord(forged, record).success).toBe(false);
    expect(writeOoxmlPackage(forged)).toEqual(snapshot);
  } finally {
    editor.destroy();
  }
}, 30000);
