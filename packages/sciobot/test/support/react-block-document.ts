import { escapeXml, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { blockTag, withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import { docxFixture } from './document';

export const hostParagraph = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
export const hostControl = (id: string, content: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
export function reactBlockDocument(prefix = ''): Uint8Array {
  const result = readOoxmlPackage(
    docxFixture(
      prefix +
        hostParagraph('Intro') +
        hostControl(
          'counter-1',
          hostParagraph('Widget fallback') + hostParagraph('Editable body')
        ) +
        hostControl('future-1', hostParagraph('Unregistered fallback')) +
        hostParagraph('Tail')
    )
  );
  if (!result.ok) throw new Error(result.reason);
  let pkg = result.package;
  for (const [id, type] of [
    ['counter-1', 'test_counter'],
    ['future-1', 'future_widget'],
  ]) {
    const written = withPayloadRecord(
      pkg,
      { id: id!, type: type!, version: 1, attributes: { name: 'Initial' } },
      'Word fallback'
    );
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return writeOoxmlPackage(pkg);
}

/** Nested blocks, native tables, row controls, and cell controls in one real DOCX. */
export function nativeSlotDocument(): Uint8Array {
  const cell = (content: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  const table = (rows: string) =>
    '<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>' +
    rows +
    '</w:tbl>';
  const body =
    hostParagraph('Intro') +
    hostControl(
      'outer',
      hostParagraph('Outer start') +
        hostControl('inner', hostParagraph('Inner text')) +
        hostParagraph('Outer middle') +
        table(
          '<w:tr>' +
            cell(
              hostControl('inside-cell', hostParagraph('Cell text')) + hostParagraph('Cell tail')
            ) +
            cell(hostParagraph('Other cell')) +
            '</w:tr>'
        ) +
        hostControl('only-child', hostControl('leaf', hostParagraph('Leaf text'))) +
        hostParagraph('Outer end')
    ) +
    table(
      hostControl(
        'row',
        '<w:tr>' +
          hostControl('cell', cell(hostParagraph('Row cell text'))) +
          cell(hostParagraph('Row other text')) +
          '</w:tr>'
      )
    ) +
    hostParagraph('Tail');
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const id of ['outer', 'inner', 'inside-cell', 'only-child', 'leaf', 'row', 'cell']) {
    const written = withPayloadRecord(
      pkg,
      {
        id,
        type: 'native_container',
        version: 1,
        attributes: { name: id },
      },
      id
    );
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return writeOoxmlPackage(pkg);
}

/** A nested native container that continues across pages, beside a form widget. */
export function blockSelectionDocument(lines = 56): Uint8Array {
  const outerLines = Array.from({ length: lines }, (_, index) =>
    hostParagraph(`Outer line ${index + 1}`)
  ).join('');
  const body =
    hostParagraph('Intro') +
    hostControl(
      'outer',
      hostParagraph('Outer start') +
        hostControl('inner', hostParagraph('Inner text')) +
        outerLines +
        hostParagraph('Outer end')
    ) +
    hostControl('form', hostParagraph('Form fallback')) +
    hostParagraph('Tail');
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const [id, type] of [
    ['outer', 'native_container'],
    ['inner', 'native_container'],
    ['form', 'test_form'],
  ] as const) {
    const written = withPayloadRecord(pkg, { id, type, version: 1, attributes: { name: id } }, id);
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return writeOoxmlPackage(pkg);
}

/** Two phases. The first holds an activity whose nested note references the activity. */
export function blockClipboardDocument(): Uint8Array {
  const body =
    hostParagraph('Intro') +
    hostControl(
      'phase-1',
      hostParagraph('Phase one') +
        hostControl(
          'activity-1',
          hostParagraph('Activity text') + hostControl('note-1', hostParagraph('Note text'))
        ) +
        hostParagraph('Phase end')
    ) +
    hostControl('phase-2', hostParagraph('Phase two')) +
    hostParagraph('Tail');
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const [id, type, attributes] of [
    ['phase-1', 'phase', { name: 'Phase one' }],
    ['activity-1', 'activity', { name: 'Warmup', duration: '10' }],
    ['note-1', 'note', { activityRef: 'activity-1', tone: 'info' }],
    ['phase-2', 'phase', { name: 'Phase two' }],
  ] as const) {
    const written = withPayloadRecord(pkg, { id, type, version: 1, attributes }, id);
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return writeOoxmlPackage(pkg);
}

/** A measured widget between native paragraphs; `rows` drives its rendered height. */
export function blockMeasurementDocument(transcript?: readonly string[]): Uint8Array {
  const body =
    hostParagraph('Intro') +
    hostControl(
      'tall',
      transcript ? transcript.map(hostParagraph).join('') : hostParagraph('Tall fallback')
    ) +
    hostParagraph('Tail') +
    hostParagraph('After');
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  const written = withPayloadRecord(
    read.package,
    { id: 'tall', type: 'test_tall', version: 1, attributes: { rows: transcript?.length ?? 0 } },
    'Tall'
  );
  if (!written.ok) throw new Error(written.reason);
  return writeOoxmlPackage(written.pkg);
}

/** A source chip between native words, with a separate block reservation fixture. */
export function inlineWidgetDocument(): Uint8Array {
  const run = (text: string) => `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  const read = readOoxmlPackage(
    docxFixture(
      `<w:p>${run('Before ')}${hostControl('source-1', run('Source'))}${run(' after')}</w:p>` +
        hostControl('tall', hostParagraph('Tall fallback')) +
        hostParagraph('Tail')
    )
  );
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const [id, type] of [
    ['source-1', 'source_chip'],
    ['tall', 'test_tall'],
  ] as const) {
    const result = withPayloadRecord(
      pkg,
      {
        id,
        type,
        version: 1,
        attributes: { name: 'Initial' },
      },
      'Source'
    );
    if (!result.ok) throw new Error(result.reason);
    pkg = result.pkg;
  }
  return writeOoxmlPackage(pkg);
}
