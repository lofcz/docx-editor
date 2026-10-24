import { expect, test } from 'bun:test';
import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { applicationFragmentOptions } from '../../test/support/application-fragments';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import type { TargetDocumentOptions } from '../diff/plan-target';
import { recomposeOwnedProposals } from './recompose';

test('working-target comparison never consumes host block identities or payload services', async () => {
  const input = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
  if (!input.ok) throw Error(input.reason);
  const editor = propertyEditor(input.package);
  try {
    const configured = applicationFragmentOptions('host-services');
    let identities = 0,
      payloads = 0;
    const target: TargetDocumentOptions = {
      ...configured,
      allocateBlockId: () => {
        identities++;
        return configured.allocateBlockId!();
      },
      newBlockPayload: (ast, id) => {
        payloads++;
        return configured.newBlockPayload!(ast, id);
      },
    };
    const markdown = 'Outside\n\n<BudgetCard title="First" />';
    const first = await recomposeOwnedProposals(
      editor,
      { revision: propertyToken(editor), targetMarkdown: markdown, ownership: [] },
      propertyContext,
      { target, payloads: { presence: { proposalId: 'host-services-presence' } } }
    );
    if (!first.success) throw Error(first.error);
    expect(first.committed).toBe(true);
    expect({ identities, payloads }).toEqual({ identities: 1, payloads: 1 });
    const before = editor.surface!.session.currentPackage(),
      token = propertyToken(editor);
    const same = await recomposeOwnedProposals(
      editor,
      {
        revision: token,
        targetMarkdown: markdown,
        ownership: first.ownership,
        contentProofs: first.contentProofs,
      },
      propertyContext,
      { target }
    );
    expect(same).toMatchObject({ success: true, committed: false });
    expect({ identities, payloads }).toEqual({ identities: 1, payloads: 1 });
    expect(editor.surface!.session.currentPackage()).toBe(before);
    expect(propertyToken(editor)).toEqual(token);
  } finally {
    editor.destroy();
  }
});
