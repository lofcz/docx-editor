import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  prepareTrackedRunOrigins,
  readOoxmlPart,
  sealTrackedRunOrigins,
  serializeOoxmlPart,
  type OoxmlPart,
} from '../index.ts';
import { ADA, apply, paragraphId, part } from './tracked-edit-fixture.ts';

function reopen(source: OoxmlPart): OoxmlPart {
  const parsed = readOoxmlPart(serializeOoxmlPart(source), {
    name: source.name,
    contentType: source.contentType,
  });
  if (!parsed.ok) throw Error(parsed.reason);
  return parsed.part;
}

for (const decision of ['accept', 'reject'] as const)
  for (const saved of [false, true])
    test(`separate native text and formatting stages settle every source scope: ${decision}/${saved}`, () => {
      const source = part('<w:p><w:r><w:t>Alpha</w:t></w:r><w:r><w:t>Bridge</w:t></w:r></w:p>');
      const id = paragraphId(source);
      const first = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 4 }]);
      let pending = apply(first.part, {
        op: 'deleteText',
        paragraphId: id,
        start: 0,
        end: 5,
        revision: ADA,
      });
      pending = apply(pending, {
        op: 'insertText',
        paragraphId: id,
        offset: 0,
        text: 'New',
        revision: ADA,
      });
      const second = prepareTrackedRunOrigins(pending, id, [{ start: 9, end: 13 }]);
      pending = apply(second.part, {
        op: 'setRunProperties',
        paragraphId: id,
        start: 9,
        end: 13,
        properties: [{ localName: 'b', attributes: {} }],
        revision: ADA,
      });
      pending = sealTrackedRunOrigins(pending, id, first.scope!, source);
      pending = sealTrackedRunOrigins(pending, id, second.scope!, source);
      if (saved) pending = reopen(pending);
      const decided = apply(pending, {
        op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
      });
      let expected = source;
      if (decision === 'accept') {
        // Authored replacement text does not donate its run identity to retained Bridge text.
        expected = part('<w:p><w:r><w:t>New</w:t></w:r><w:r><w:t>Bridge</w:t></w:r></w:p>');
        expected = apply(expected, {
          op: 'setRunProperties',
          paragraphId: id,
          start: 4,
          end: 8,
          properties: [{ localName: 'b', attributes: {} }],
        });
      }
      expect(canonicalOoxmlFingerprint(decided)).toBe(canonicalOoxmlFingerprint(expected));
      expect(canonicalOoxmlFingerprint(reopen(decided))).toBe(canonicalOoxmlFingerprint(expected));
    });
