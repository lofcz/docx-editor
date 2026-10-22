import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { blankDocumentBytes } from '../blank-document.ts';
import {
  collectReviewItems,
  collectRevisionSites,
  customXmlDataParts,
  customXmlNodes,
  findNode,
  paragraphTextOf,
  revisionItemsOf,
  withCustomXmlDataPart,
  withCustomXmlNode,
  writeOoxmlPackage,
  type OoxmlPackage,
  type OoxmlPart,
} from '../../store/index.ts';

const NS = 'urn:test:post-ops',
  module = {
    id: 'test-post-ops-review',
    review: {
      collectReviewItems,
      displayModes: ['all-markup', 'simple-markup', 'proposed', 'original'] as const,
      revisionItemsOfParagraph: (part: OoxmlPart, id: string) => {
        const node = findNode(part, id);
        return node?.kind === 'paragraph' ? revisionItemsOf({ ...part, root: node }) : [];
      },
    },
  };
function mount() {
  return createDocxEditor({
    container: document.createElement('div'),
    document: blankDocumentBytes(),
    modules: [module],
  });
}
function receipt(pkg: OoxmlPackage) {
  const part = customXmlDataParts(pkg, pkg.mainDocumentPart).find(
    (entry) => entry.namespaceUri === NS
  );
  return part ? customXmlNodes(pkg, part.partName)[0]?.data : undefined;
}
function withReceipt(pkg: OoxmlPackage, value: string) {
  const created = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, NS, 'receipts');
  if (!created.part) throw new Error('Missing native receipt part');
  return withCustomXmlNode(created.pkg, created.part.partName, {
    id: 'receipt',
    label: '',
    data: value,
  });
}

test('post hooks read actual tracked identities and publish metadata with text in one undo unit', () => {
  const editor = mount();
  try {
    const surface = editor.surface!,
      session = surface.session,
      id = session.paragraphIds()[0]!,
      before = session.packageRevision(),
      publications: [string | null, string | undefined][] = [],
      stop = session.subscribe(() =>
        publications.push([paragraphTextOf(session.part(), id), receipt(session.currentPackage())])
      );
    const result = surface.applyAutomationOps(
      () => [
        {
          op: 'insertText',
          paragraphId: id,
          offset: 0,
          text: 'New',
          revision: { author: 'ScioBot', date: '2026-10-19T09:00:00.000Z' },
        },
      ],
      { kind: 'body' },
      [
        (pkg) => {
          const sites = collectRevisionSites(pkg.parts.get(pkg.mainDocumentPart)!);
          expect(sites).toHaveLength(1);
          return withReceipt(pkg, JSON.stringify(sites[0]!.node.attributes));
        },
      ],
      true,
      true,
      'after'
    );
    stop();
    expect(result.committed).toBe(true);
    expect(publications).toHaveLength(1);
    expect(publications[0]![0]).toBe('New');
    expect(publications[0]![1]).toBe(
      JSON.stringify(collectRevisionSites(session.part())[0]!.node.attributes)
    );
    expect(session.packageRevision()).toBe(before + 1);
    const pending = writeOoxmlPackage(session.currentPackage());
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('');
    expect(receipt(session.currentPackage())).toBeUndefined();
    expect(session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(pending);
  } finally {
    editor.destroy();
  }
});

test('post hook refusal rolls back native edits, earlier hook metadata, revision, and history', () => {
  const editor = mount();
  try {
    const surface = editor.surface!,
      session = surface.session,
      id = session.paragraphIds()[0]!,
      bytes = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision();
    let publications = 0;
    const stop = session.subscribe(() => publications++);
    expect(() =>
      surface.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: id, offset: 0, text: 'Must roll back' }],
        { kind: 'body' },
        [
          (pkg) => withReceipt(pkg, 'rollback'),
          () => {
            throw new Error('post-hook-refusal');
          },
        ],
        false,
        true,
        'after'
      )
    ).toThrow('post-hook-refusal');
    stop();
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(revision);
    expect(publications).toBe(0);
    expect([session.canUndo(), session.canRedo()]).toEqual([false, false]);
  } finally {
    editor.destroy();
  }
});

test('before hooks keep their existing order while native failures and viewing skip post hooks', () => {
  const editor = mount();
  try {
    const surface = editor.surface!,
      session = surface.session,
      id = session.paragraphIds()[0]!;
    let called = 0;
    expect(
      surface.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: id, offset: 0, text: 'A' }],
        { kind: 'body' },
        [
          (pkg) => {
            expect(paragraphTextOf(pkg.parts.get(pkg.mainDocumentPart)!, id)).toBe('');
            return pkg;
          },
        ]
      ).committed
    ).toBe(true);
    const post = (pkg: OoxmlPackage) => {
      called++;
      return pkg;
    };
    expect(
      surface.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: 'missing', offset: 0, text: 'No' }],
        { kind: 'body' },
        [post],
        false,
        true,
        'after'
      ).rejected
    ).toBe(true);
    surface.setEditingMode('view');
    expect(
      surface.applyAutomationOps(() => [], { kind: 'body' }, [post], false, true, 'after').rejected
    ).toBe(true);
    expect(called).toBe(0);
    expect(paragraphTextOf(session.part(), id)).toBe('A');
  } finally {
    editor.destroy();
  }
});

test('atomic session groups run post hooks after all operations and roll back together on failure', () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      id = session.paragraphIds()[0]!,
      historyToken = Symbol('atomic-post');
    const result = session.applyTreeOpsAtomic(
      [
        {
          scope: { kind: 'body' },
          ops: [{ op: 'insertText', paragraphId: id, offset: 0, text: 'Native' }],
        },
      ],
      {
        historyToken,
        packageEditsAfterOps: [
          (pkg) => {
            expect(paragraphTextOf(pkg.parts.get(pkg.mainDocumentPart)!, id)).toBe('Native');
            return withReceipt(pkg, 'all-groups');
          },
        ],
      }
    );
    expect(result.committed).toBe(true);
    const bytes = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision();
    expect(() =>
      session.applyTreeOpsAtomic(
        [
          {
            scope: { kind: 'body' },
            ops: [{ op: 'insertText', paragraphId: id, offset: 6, text: ' Failure' }],
          },
        ],
        {
          historyToken: Symbol('atomic-refusal'),
          packageEditsAfterOps: [
            () => {
              throw new Error('atomic-post-refusal');
            },
          ],
        }
      )
    ).toThrow('atomic-post-refusal');
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(revision);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('');
    expect(receipt(session.currentPackage())).toBeUndefined();
  } finally {
    editor.destroy();
  }
});
