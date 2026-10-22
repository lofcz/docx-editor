import { expect } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { customXmlNodes, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { readPayloadStore, withPayloadRecord } from '../../src/blocks/payload-opc-codec';

export const labelSource = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
export function labelStore(pkg: OoxmlPackage) {
  const result = readPayloadStore(pkg);
  if (result.status !== 'ready')
    throw new Error(result.status === 'invalid' ? result.reason : result.status);
  return result;
}
export function labelEntry(pkg: OoxmlPackage) {
  return labelStore(pkg).entries.get('activity')!;
}
export function labelPayload(pkg: OoxmlPackage) {
  const decoded = labelEntry(pkg).decoded;
  if (decoded.status !== 'supported') throw new Error('Unsupported label owner');
  return decoded.value;
}
export function legacyLabelNode(pkg: OoxmlPackage) {
  return customXmlNodes(pkg, labelStore(pkg).part.partName).find(
    (node) => node.id === 'scio_activity'
  )!;
}
export function writeLabel(
  editor: DocxEditorInstance,
  label: string,
  attribute: 'duration' | 'name',
  value: number | string
) {
  const payload = labelPayload(labelSource(editor));
  const applied = editor.surface!.applyAutomationOps(
    () => [],
    { kind: 'body' },
    [
      (pkg) => {
        const result = withPayloadRecord(
          pkg,
          { ...payload, attributes: { ...payload.attributes, [attribute]: value } },
          label
        );
        if (!result.ok) throw new Error(result.reason);
        return result.pkg;
      },
    ],
    false,
    true
  );
  expect(applied.rejected).toBe(false);
  expect(applied.committed).toBe(true);
}
