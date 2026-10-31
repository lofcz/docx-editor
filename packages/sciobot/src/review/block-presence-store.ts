import { activeStorage } from '../registry/active';
import { customXmlNodes, withCustomXmlNode, type OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { readPropertyReviewStore } from './property-record-store';
import { bindBlockPresenceRecord } from './block-presence-plan';
import {
  encodeBlockPresenceRecord,
  decodeBlockPresenceRecord,
  type BlockPresenceRecord,
} from './block-presence-schema';

/** Dormant native storage primitive. Review-group admission must own the eventual live transaction. */
export function withBlockPresenceRecord(
  pkg: OoxmlPackage,
  record: BlockPresenceRecord
): { readonly success: true; readonly pkg: OoxmlPackage } | NativeEditFailure {
  const bound = bindBlockPresenceRecord(pkg, record);
  if (!bound.success) return bound;
  const store = readPropertyReviewStore(pkg);
  if (store.status !== 'ready')
    return {
      success: false,
      code: 'invalid-document',
      error: 'Presence records require the existing application metadata store.',
    };
  const raw = encodeBlockPresenceRecord(bound.record),
    old = store.retainedRecords.get(record.id);
  if (
    old &&
    customXmlNodes(pkg, store.part.partName).find(
      (node) => node.id === activeStorage().reviewNodePrefix + record.id
    )?.label !== ''
  )
    return {
      success: false,
      code: 'stale-revision',
      error: 'A resolved presence proposal identity stays reserved.',
    };
  if (old)
    return old.raw === raw
      ? { success: true, pkg }
      : {
          success: false,
          code: 'stale-revision',
          error: 'A presence proposal identity cannot be rewritten.',
        };
  let bytes = new TextEncoder().encode(raw).byteLength;
  for (const entry of store.retainedRecords.values())
    bytes += new TextEncoder().encode(entry.raw).byteLength;
  const root = pkg.parts.get(store.part.partName)!.root;
  if (
    store.retainedRecords.size >= 1024 ||
    root.children.length >= 50000 ||
    bytes > 8 * 1024 * 1024
  )
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'Presence storage exceeds the existing review metadata limits.',
    };
  const next = withCustomXmlNode(pkg, store.part.partName, {
    id: activeStorage().reviewNodePrefix + record.id,
    label: '',
    data: raw,
  });
  const read = readPropertyReviewStore(next),
    saved = read.status === 'ready' ? read.records.get(record.id) : undefined;
  const checked = saved && decodeBlockPresenceRecord(saved.raw);
  if (read.status !== 'ready' || saved?.raw !== raw || checked?.status !== 'supported')
    return {
      success: false,
      code: 'apply-failed',
      error: 'Native presence storage refused the record.',
    };
  return { success: true, pkg: next };
}
