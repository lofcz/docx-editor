import {
  ATOMIC_CONTENT_NAMESPACE,
  applyTreeOp,
  contentControlPropertiesOf,
  findNode,
  isAtomicContentControl,
  replaceNode,
  withAtomicContentMarker,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from './block-index';
import { convertMathControl } from './math-convert';
import { decodeMathPayload } from './math-codec';
import { encodeBlockPayload } from './payload-schema';

/** Opt in only durable controls backed by a supported application equation payload. */
export function withAtomicMathFallbacks(pkg: OoxmlPackage): OoxmlPackage {
  let next = pkg;
  for (const block of buildBlockIndex(pkg).byId.values()) {
    const math = decodeMathPayload(encodeBlockPayload(block.payload));
    if (math.status !== 'supported') continue;
    let part = next.parts.get(block.partName)!;
    let control = findNode(part, block.controlId);
    if (control?.kind !== 'contentControl') throw Error('The equation control is missing.');
    const marker = control.attributes.find(
      (attribute) =>
        attribute.namespaceUri === ATOMIC_CONTENT_NAMESPACE &&
        attribute.localName === 'atomicContent'
    );
    if (marker && marker.value !== '1')
      throw Error('The equation uses an unsupported fallback ownership version.');
    if (isAtomicContentControl(control)) continue;
    // Normalize only validated application equations before room seeding.
    next = convertMathControl(next, block.partName, math.payload);
    part = next.parts.get(block.partName)!;
    control = findNode(part, block.controlId)!;
    const lock = contentControlPropertiesOf(control).lock;
    if (lock === 'unlocked' || lock === 'sdtLocked') {
      const locked = applyTreeOp(part, {
        op: 'setContentControlProperties',
        controlId: control.id,
        lock: lock === 'sdtLocked' ? 'sdtContentLocked' : 'contentLocked',
      });
      if (!locked.ok) throw Error('The equation fallback cannot be locked.');
      part = locked.part;
      control = findNode(part, block.controlId)!;
    }
    if (control.kind !== 'contentControl') throw Error('The equation control is invalid.');
    const marked = replaceNode(part, control.id, withAtomicContentMarker(control));
    if (!marked.ok) throw Error('The equation ownership marker is invalid.');
    next = withPart(next, marked.part);
  }
  return next;
}
