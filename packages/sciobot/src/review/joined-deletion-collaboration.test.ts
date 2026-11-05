// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { serializeOoxmlPart, type OoxmlNode } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { reviewModule } from '../../../pro/src/review/review-module';

const children = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
function named(node: OoxmlNode, kind: OoxmlNode['kind']): OoxmlNode[] {
  return [...(node.kind === kind ? [node] : []), ...children(node).flatMap((n) => named(n, kind))];
}
const deletion = (text: string) =>
  `<w:del w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r>${text}</w:r></w:del>`;

for (const removal of ['link', 'control'] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    test(`joined deletion keeps typing and peer deletion through history (${removal}, ${clients})`, async () => {
      const middle = deletion('<w:delText>r</w:delText>');
      const wrapper =
        removal === 'link'
          ? `<w:hyperlink w:anchor="a">${middle}</w:hyperlink>`
          : `<w:sdt><w:sdtPr/><w:sdtContent>${middle}</w:sdtContent></w:sdt>`;
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        editor: { modules: [reviewModule()] },
        document: docxFixture(
          '<w:p><w:r><w:t>a</w:t></w:r>' +
            deletion('<w:delText xml:space="preserve">the </w:delText>') +
            wrapper +
            deletion('<w:delText>est</w:delText>') +
            '<w:r><w:t>b</w:t></w:r></w:p>'
        ),
      });
      try {
        const [alice, bob] = pair.editors;
        const a = alice!.surface!,
          b = bob!.surface!;
        const paragraphId = named(a.session.part().root, 'paragraph')[0]!.id;
        const holder = named(
          a.session.part().root,
          removal === 'link' ? 'hyperlink' : 'contentControl'
        )[0]!.id;
        pair.pause();
        expect(
          a.applyAutomationOps(() => [
            removal === 'link'
              ? { op: 'removeHyperlink', linkId: holder }
              : { op: 'removeContentControl', controlId: holder },
          ]).rejected
        ).toBe(false);
        expect(
          b.applyAutomationOps(() => [{ op: 'deleteText', paragraphId, start: 9, end: 10 }])
            .rejected
        ).toBe(false);
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        expect(
          a.applyAutomationOps(
            () => [
              {
                op: 'insertText',
                paragraphId,
                offset: 4,
                text: 'X',
                revision: { author: 'Alice', date: '2026-01-03T00:00:00Z' },
              },
            ],
            undefined,
            undefined,
            true,
            true
          )
        ).toMatchObject({ rejected: false });
        await pair.settle();
        for (const editor of pair.editors) {
          const part = editor.surface!.session.part();
          expect(named(part.root, 'revisionDelete')).toHaveLength(1);
          expect(serializeOoxmlPart(part)).toMatch(/>the rest<\/w:delText><\/w:r><\/w:del><w:ins /);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(named(a.session.part().root, 'revisionInsert')).toHaveLength(0);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await pair.savedDigests()).toEqual(saved);
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await alice!.save()),
        });
        try {
          expect(await reopened.savedDigests()).toEqual(saved);
        } finally {
          reopened.destroy();
        }
      } finally {
        pair.destroy();
      }
    });
