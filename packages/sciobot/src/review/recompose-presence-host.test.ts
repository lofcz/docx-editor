import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';

const context = {
  documentId: 'presence-host',
  epoch: 'load',
  author: 'AI',
  date: '2026-10-26T12:00:00Z',
};
const target = 'Start\n\n| A | B |\n| - | - |\n| $x$ [Ada](mention:user) | Tail |';
const create = () =>
  createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'),
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });

test('the host allocates presence identities only for application targets across replacement waves', async () => {
  const editor = create();
  const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
  let calls = 0;
  const options = { createPresenceProposalId: () => `host-presence-${++calls}` };
  try {
    const plain = await recomposeOwnedProposals(
      editor,
      { revision: token(), targetMarkdown: 'Start', ownership: [] },
      context,
      options
    );
    expect(plain).toMatchObject({ success: true, committed: false });
    expect(calls).toBe(0);
    const first = await recomposeOwnedProposals(
      editor,
      { revision: token(), targetMarkdown: target, ownership: [] },
      context,
      options
    );
    if (!first.success) throw Error(first.error);
    expect(calls).toBe(1);
    const next = target.replace('$x$', '$y$').replace('Ada', 'Grace');
    const second = await recomposeOwnedProposals(
      editor,
      {
        revision: token(),
        targetMarkdown: next,
        ownership: first.ownership,
        contentProofs: first.contentProofs,
      },
      context,
      options
    );
    if (!second.success) throw Error(second.error);
    expect(calls).toBe(2);
    const working = await readWorkingView(editor.surface!.session, token());
    if (!working.success) throw Error(working.error);
    expect(working.view.projection.text).toContain('Grace');
    const current = editor.surface!.session.currentPackage();
    const repeated = await recomposeOwnedProposals(
      editor,
      {
        revision: token(),
        targetMarkdown: next,
        ownership: second.ownership,
        contentProofs: second.contentProofs,
      },
      context,
      options
    );
    expect(repeated).toMatchObject({ success: true, committed: false });
    expect(calls).toBe(2);
    expect(editor.surface!.session.currentPackage()).toBe(current);
    const reopened = readOoxmlPackage(writeOoxmlPackage(current));
    if (!reopened.ok) throw Error(reopened.reason);
    expect(
      [...buildBlockIndex(reopened.package).byId.values()].map((block) => block.payload.type)
    ).toEqual(['inline_equation', 'mention']);
  } finally {
    editor.destroy();
  }
});

for (const failure of ['throw', 'empty'] as const)
  test(`a ${failure} host presence allocation leaves native content and history unchanged`, async () => {
    const editor = create();
    try {
      const before = editor.surface!.session.currentPackage();
      const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const result = await recomposeOwnedProposals(
        editor,
        { revision, targetMarkdown: target, ownership: [] },
        context,
        {
          createPresenceProposalId: () => {
            if (failure === 'throw') throw Error('Host allocation failed.');
            return '';
          },
        }
      );
      expect(result.success).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
        canonicalOoxmlFingerprint(before.parts.get(before.mainDocumentPart)!)
      );
      expect(editor.surface!.session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

test('host allocation cannot commit an AI target after an intervening native edit', async () => {
  const editor = create();
  try {
    const surface = editor.surface!;
    const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
    const result = await recomposeOwnedProposals(
      editor,
      { revision, targetMarkdown: target, ownership: [] },
      context,
      {
        createPresenceProposalId: () => {
          const paragraphId = surface.session.paragraphIds()[0]!;
          const applied = surface.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset: 0, text: 'Peer ' },
          ]);
          expect(applied).toMatchObject({ committed: true, rejected: false });
          return 'stale-host-presence';
        },
      }
    );
    expect(result).toMatchObject({ success: false, code: 'stale-revision' });
    expect(buildBlockIndex(surface.session.currentPackage()).byId.size).toBe(0);
    const current = readEditorRevision(editor, context.documentId, context.epoch)!;
    const working = await readWorkingView(surface.session, current);
    if (!working.success) throw Error(working.error);
    expect(working.view.projection.text).toBe('Peer Start');
  } finally {
    editor.destroy();
  }
});
