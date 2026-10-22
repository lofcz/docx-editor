import { escapeXml, readXml, type OoxmlElement, type OoxmlNode } from '@docx-editor.dev/core/store';

/** Each authored update owns one immutable scalar value and one independent history identity. */
export const PAYLOAD_LABEL_ELEMENT = 'labelValue';
export const PAYLOAD_LABEL_LIMIT = 2 * 1024 * 1024;
export const PAYLOAD_LABEL_RECORD_LIMITS = Object.freeze({
  maxRecords: 10000,
  maxSlots: 50000,
  maxBytes: 8 * 1024 * 1024,
});
export const isPayloadLabel = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length <= PAYLOAD_LABEL_LIMIT &&
  readXml(`<label>${escapeXml(value)}</label>`).ok;
const isRecord = (node: OoxmlNode, namespaceUri: string): node is OoxmlElement =>
  node.kind !== 'textValue' &&
  node.namespaceUri === namespaceUri &&
  node.localName === PAYLOAD_LABEL_ELEMENT;

export function readPayloadLabel(
  record: OoxmlElement | undefined,
  legacy: string
): { readonly ok: true; readonly label: string } | { readonly ok: false; readonly reason: string } {
  if (record && record.children.length > PAYLOAD_LABEL_RECORD_LIMITS.maxSlots)
    return { ok: false, reason: 'label-register-slot-limit' };
  const records = record?.children.filter((node) => isRecord(node, record.namespaceUri)) ?? [];
  let count = 0,
    bytes = 0,
    label = legacy;
  for (const child of records) {
    if (!isRecord(child, record!.namespaceUri)) continue;
    // Native undo can retain an anonymous empty wrapper; it carries no authored value.
    if (!child.attributes.length && !child.children.length) continue;
    if (++count > PAYLOAD_LABEL_RECORD_LIMITS.maxRecords)
      return { ok: false, reason: 'label-register-count-limit' };
    const version = child.attributes.filter(
        (a) => a.namespaceUri === '' && a.localName === 'version'
      ),
      values = child.attributes.filter((a) => a.namespaceUri === '' && a.localName === 'value');
    if (version.length !== 1 || version[0]!.value !== '1')
      return { ok: false, reason: 'unsupported-label-register-version' };
    if (values.length !== 1 || child.children.length || child.attributes.length !== 2)
      return { ok: false, reason: 'invalid-label-register' };
    const raw = values[0]!.value;
    if (raw.length > PAYLOAD_LABEL_LIMIT * 2 + 2)
      return { ok: false, reason: 'label-register-limit' };
    bytes += new TextEncoder().encode(raw).byteLength;
    if (bytes > PAYLOAD_LABEL_RECORD_LIMITS.maxBytes)
      return { ok: false, reason: 'label-register-byte-limit' };
    let value: unknown;
    try {
      value = JSON.parse(raw);
    } catch {
      return { ok: false, reason: 'invalid-label-register' };
    }
    if (!isPayloadLabel(value)) return { ok: false, reason: 'invalid-label-register' };
    label = value;
  }
  return isPayloadLabel(label) ? { ok: true, label } : { ok: false, reason: 'invalid-label' };
}

/** Append instead of replacing another actor's register. Document order selects one whole value. */
export function withPayloadLabel(record: OoxmlElement, label: string): OoxmlElement {
  if (!isPayloadLabel(label)) throw new TypeError('invalid-label');
  const current = readPayloadLabel(record, '');
  if (!current.ok) throw new TypeError(current.reason);
  if (
    record.children.some(
      (node) =>
        isRecord(node, record.namespaceUri) && (node.attributes.length || node.children.length)
    ) &&
    current.label === label
  )
    return record;
  const next: OoxmlElement = {
    id: `label_${crypto.randomUUID().replaceAll('-', '')}`,
    kind: 'generic',
    namespaceUri: record.namespaceUri,
    localName: PAYLOAD_LABEL_ELEMENT,
    namespaceBindings: [],
    attributes: [
      { kind: 'genericExtension', namespaceUri: '', localName: 'version', value: '1' },
      {
        kind: 'genericExtension',
        namespaceUri: '',
        localName: 'value',
        value: JSON.stringify(label),
      },
    ],
    children: [],
  };
  return { ...record, children: [...record.children, next] } as OoxmlElement;
}
