import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, collectRevisionSites } from '../index';
import { runWithTransactionActor } from '../package/actor-scoped-ids';
import { ADA, apply, part, paragraphId, xml } from './tracked-edit-fixture';

for (const actor of [undefined, 'alice'])
  test(`accepted whole-table insertion matches direct native structure; actor=${actor}`, () => {
    const before = part('<w:p><w:r><w:t>After</w:t></w:r></w:p>');
    const operation = {
      op: 'insertTable',
      beforeParagraphId: paragraphId(before),
      rows: 2,
      cols: 2,
      columnWidthTwips: 2000,
    } as const;
    const direct = apply(before, operation);
    const tracked = runWithTransactionActor(actor, () =>
      apply(before, { ...operation, revision: ADA })
    );
    const accepted = runWithTransactionActor(actor, () =>
      apply(tracked, { op: 'acceptAllRevisions' })
    );
    expect(collectRevisionSites(accepted)).toHaveLength(0);
    expect(canonicalOoxmlFingerprint(accepted)).toBe(canonicalOoxmlFingerprint(direct));
    expect(xml(accepted)).not.toContain('<w:trPr/>');
  });

for (const extra of [' x:stamp="keep"', ''])
  test(`accepted row and cell holders retain opaque properties; attributes=${Boolean(extra)}`, () => {
    const mark = 'w:id="1" w:author="Ada Lovelace"';
    const rowExtra = extra ? '' : '<x:rowData/>';
    const cellExtra = extra ? '' : '<x:cellData/>';
    const source = part(
      `<w:tbl xmlns:x="urn:keep"><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr${extra}><w:ins ${mark}/>${rowExtra}</w:trPr><w:tc><w:tcPr${extra}><w:cellIns ${mark}/>${cellExtra}</w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/>`
    );
    const accepted = apply(source, { op: 'acceptAllRevisions' });
    expect(collectRevisionSites(accepted)).toHaveLength(0);
    expect(xml(accepted)).toContain(extra ? 'x:stamp="keep"' : '<x:rowData/>');
    expect(xml(accepted)).toContain(extra ? '<w:tcPr x:stamp="keep"' : '<x:cellData/>');
  });
