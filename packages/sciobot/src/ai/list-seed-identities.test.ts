import { expect, test } from 'bun:test';
import {
  readOoxmlPackage,
  contentControlsIn,
  serializeOoxmlPart,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { propertyEditor } from '../../test/support/property-execution';
import { planListMembership } from '../diff/lists';
import { readEditorRevision } from '../projection/revision-cache';
import { applyDirectPlan } from './apply-direct';
const context = { documentId: 'list-seed-identities', epoch: 'mount' };
for (const checklist of [false, true])
  test(`separate ${checklist ? 'checklist creations' : 'ordered starts'} preserve unique canonical identities`, async () => {
    const read = readOoxmlPackage(
      docxFixture('<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>')
    );
    if (!read.ok) throw Error(read.reason);
    const editor = propertyEditor(read.package);
    try {
      const session = editor.surface!.session,
        ids = [...session.paragraphIds()];
      for (const [index, paragraphId] of ids.entries()) {
        const planned = planListMembership(
          {
            pkg: session.currentPackage(),
            partName: session.part().name,
            revision: readEditorRevision(editor, context.documentId, context.epoch)!,
          },
          [
            {
              paragraphId,
              target: checklist
                ? { kind: 'checklist', level: 0, checked: index === 1 }
                : {
                    kind: 'ordered',
                    level: 0,
                    instance: { kind: 'new', key: `instance-${index}`, start: 5 + index * 2 },
                  },
            },
          ]
        );
        if (!planned.success) throw Error(planned.error);
        const applied = applyDirectPlan(editor, planned.plan, context);
        if (!applied.success) throw Error(applied.error);
      }
      const canonical = new Set<string>();
      let count = 0;
      const visit = (node: OoxmlNode) => {
        count++;
        canonical.add(node.id);
        if (node.kind !== 'textValue') node.children.forEach(visit);
      };
      session.currentPackage().parts.forEach((part) => visit(part.root));
      expect(canonical.size).toBe(count);
      const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw Error(saved.reason);
      if (checklist) {
        const controls = contentControlsIn(
          saved.package.parts.get(saved.package.mainDocumentPart)!.root
        );
        expect(controls).toHaveLength(2);
        expect(new Set(controls.map((control) => control.node.id)).size).toBe(2);
      } else {
        const numbering = saved.package.parts.get('/word/numbering.xml');
        expect(numbering).toBeDefined();
        const xml = serializeOoxmlPart(numbering!);
        expect(xml).toContain('<w:startOverride w:val="5"/>');
        expect(xml).toContain('<w:startOverride w:val="7"/>');
        expect(xml.match(/<w:num w:numId=/g)).toHaveLength(2);
      }
    } finally {
      editor.destroy();
    }
  });
