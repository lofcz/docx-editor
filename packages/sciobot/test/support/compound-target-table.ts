import { docxFixture } from './document';
export interface TargetCell {
  readonly text: string;
  readonly columns?: number;
  readonly rows?: number;
}
export interface CompoundTableCase {
  readonly name: string;
  readonly source?: string;
  readonly rows: readonly (readonly TargetCell[])[];
}
export const cell = (text: string, columns = 1, rows = 1): TargetCell => ({ text, columns, rows });
const p = (text: string) =>
  `<w:p><w:pPr><w:rPr/></w:pPr><w:r><w:rPr/><w:t>${text}</w:t></w:r></w:p>`;
const tc = (text: string, span = 1) =>
  `<w:tc><w:tcPr><w:tcW w:w="${2400 * span}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}</w:tcPr>${p(text)}</w:tc>`;
export const document_ = (merged = false) =>
  docxFixture(
    p('Lead') +
      '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>' +
      '<w:gridCol w:w="2400"/>'.repeat(3) +
      '</w:tblGrid>' +
      (merged
        ? '<w:tr><w:trPr/>' + tc('H1', 2) + tc('H3') + '</w:tr>'
        : '<w:tr><w:trPr/>' + ['H1', 'H2', 'H3'].map((t) => tc(t)).join('') + '</w:tr>') +
      [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ]
        .map((row) => '<w:tr><w:trPr/>' + row.map((t) => tc(t)).join('') + '</w:tr>')
        .join('') +
      '</w:tbl>' +
      p('Tail')
  );
export const target = (rows: readonly (readonly TargetCell[])[]) =>
  'Lead\n\n<table>' +
  rows
    .map(
      (row) =>
        '<tr>' +
        row
          .map(
            (c) =>
              `<td${c.columns && c.columns > 1 ? ` colspan="${c.columns}"` : ''}${c.rows && c.rows > 1 ? ` rowspan="${c.rows}"` : ''}>\n\n${c.text}\n\n</td>`
          )
          .join('') +
        '</tr>'
    )
    .join('') +
  '</table>\n\nTail';
const original = () =>
  [
    ['H1', 'H2', 'H3'],
    ['A', 'B', 'C'],
    ['D', 'E', 'F'],
  ].map((row) => row.map((t) => cell(t)));
export const cases: CompoundTableCase[] = [];
for (const at of [0, 1, 3]) {
  const rows = original();
  rows.splice(at, 0, [cell('New left\n\nNew right', 2), cell('New tail')]);
  cases.push({ name: `new row ${at} merges allocated cells`, rows });
}
for (const at of [0, 1, 3]) {
  const rows = original().map((row, r) => {
    row.splice(at, 0, cell(`New ${r}`));
    if (r === 0) {
      const first = at === 3 ? 2 : at;
      const combined = row[first]!.text + '\n\n' + row[first + 1]!.text;
      row.splice(first, 2, cell(combined, 2));
    }
    return row;
  });
  cases.push({ name: `new column ${at} merges allocated and retained cells`, rows });
}
{
  const rows = original();
  rows[0] = [cell('H1\n\nH2', 2), cell('H3'), cell('N0')];
  rows[1]!.push(cell('N1'));
  rows[2]!.push(cell('N2'));
  rows.push([cell('New 1\n\nNew 2', 2), cell('New 3'), cell('New 4')]);
  cases.push({ name: 'row and column growth creates two disjoint merges', rows });
}
cases.push({
  name: 'column deletion precedes retained merge',
  rows: [[cell('H1\n\nH3', 2)], [cell('A'), cell('C')], [cell('D'), cell('F')]],
});
cases.push({
  name: 'row deletion rebases a vertical merge',
  rows: [
    [cell('H1\n\nD', 1, 2), cell('H2'), cell('H3')],
    [cell('E'), cell('F')],
  ],
});
cases.push({
  name: 'new first row starts a vertical merge with retained content',
  rows: [
    [cell('New\n\nH1', 1, 2), cell('New 2'), cell('New 3')],
    [cell('H2'), cell('H3')],
    [cell('A'), cell('B'), cell('C')],
    [cell('D'), cell('E'), cell('F')],
  ],
});
cases.push({
  name: 'split retained span after new row allocation',
  source: 'merged',
  rows: [
    [cell('H1'), cell('H2'), cell('H3')],
    [cell('A'), cell('B'), cell('C')],
    [cell('D'), cell('E'), cell('F')],
    [cell('New 1'), cell('New 2'), cell('New 3')],
  ],
});
cases.push({
  name: 'new row inherits and then splits its anchor span',
  source: 'merged',
  rows: [
    [cell('New 1'), cell('New 2'), cell('New 3')],
    [cell('H1', 2), cell('H3')],
    [cell('A'), cell('B'), cell('C')],
    [cell('D'), cell('E'), cell('F')],
  ],
});
cases.push({
  name: 'column inserted inside a retained span then split',
  source: 'merged',
  rows: [
    [cell('H1'), cell('Added'), cell('H2'), cell('H3')],
    [cell('A'), cell('N1'), cell('B'), cell('C')],
    [cell('D'), cell('N2'), cell('E'), cell('F')],
  ],
});
