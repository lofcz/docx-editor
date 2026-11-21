import { afterEach, expect, test } from 'bun:test';
import { createDocxEditor, type SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  contentControlTextOf,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  storyParagraphs,
  TreeDocumentStore,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  configureMdxDisclosures,
  createMdxComponentRegistry,
  exportMdxDocument,
  importMdxDocument,
  mountMdxComponents,
  mountMdxDisclosures,
  readMdxComponents,
  type BlockPayload,
  type MdxComponentParser,
  type MdxComponentView,
} from '../src/mdx.ts';
import { parseDocumentSyntax } from '../src/mdx/document-syntax.ts';

function project(slug: string, units: string) {
  const read: MdxComponentParser['read'] = (attrs) => {
    const reading = attrs.find((attr) => attr.name === 'reading')?.value;
    return typeof reading === 'string' ? { reading } : null;
  };
  const result = createMdxComponentRegistry({
    storage: {
      tagPrefix: slug + ':1:',
      namespace: 'urn:' + slug + ':records',
      root: 'records',
      nodePrefix: slug + '_',
      reviewNodePrefix: slug + '-review_',
      labelTag: slug + ':label',
      adapterPrefix: slug,
      adapterNamespace: 'urn:' + slug + ':word',
    },
    components: [
      {
        type: 'measurement',
        kind: 'block',
        content: 'none',
        mdx: { tag: 'Reading', kind: 'atomic', attributes: () => [] },
        parse: { tag: 'Reading', read },
        wordFallback: ({ attributes }) => ({
          lines: [slug + ' instrument', attributes.reading + ' ' + units],
        }),
      },
      {
        type: 'stamp',
        kind: 'inline',
        content: 'none',
        mdx: { tag: 'Stamp', kind: 'atomic', attributes: () => [] },
        parse: { tag: 'Stamp', read },
        inline: {
          atom: 'Stamp',
          fallback: (payload) => ({ payload, text: units + ' ' + payload.attributes.reading }),
        },
      },
    ],
  });
  if (!result.ok) throw new Error(result.detail);
  return result.registry;
}

function open(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}

function wordParagraphs(pkg: OoxmlPackage) {
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  return storyParagraphs(bodyStoryRoot(part)!).map(({ id }) => ({
    id,
    text: paragraphTextOf(part, id),
  }));
}

const instruments = '<Reading reading="12" />\n\nSee <Stamp reading="7" /> below.';

test('registered math fallbacks preserve formula source inside native lists and table cells', () => {
  const payload = (
    type: string,
    id: string,
    value: BlockPayload['attributes']['value']
  ): BlockPayload => ({
    version: 1,
    id,
    type,
    attributes: { value },
  });
  const result = createMdxComponentRegistry({
    storage: project('lab', 'mm').storage,
    components: [
      {
        type: 'display',
        kind: 'block',
        content: 'none',
        markdown: {
          types: ['math'],
          fromMarkdown: (node, id) => payload('display', id, node.value),
          matches: () => true,
        },
        wordFallback: ({ attributes }) => ({ lines: ['Equation: ' + attributes.value] }),
      },
      {
        type: 'formula',
        kind: 'inline',
        content: 'none',
        inline: {
          atom: 'formula',
          types: ['inlineMath'],
          fromMarkdown: (node, id) => payload('formula', id, node.value),
          fallback: (data) => ({ payload: data, text: 'Variable ' + data.attributes.value }),
        },
      },
    ],
  });
  if (!result.ok) throw new Error(result.detail);
  const source = '<Panel>\n\n- Read $v$.\n\n| Value |\n| --- |\n| $p$ |\n\n$$\nx+y\n$$\n\n</Panel>';
  const pkg = open(importMdxDocument(source, { registry: result.registry }));
  const paragraphs = wordParagraphs(pkg);
  expect(paragraphs.map((p) => p.text)).toEqual([
    'Read Variable v.',
    'Value',
    'Variable p',
    'Equation: x+y',
  ]);
  expect(parseDocumentSyntax(exportMdxDocument(writeOoxmlPackage(pkg)))).toEqual(
    parseDocumentSyntax(source)
  );
  const store = new TreeDocumentStore(pkg.parts.get(pkg.mainDocumentPart)!);
  expect(
    store.transact((ctx) =>
      ctx.apply({
        op: 'insertText',
        paragraphId: paragraphs.at(-1)!.id,
        offset: 2,
        text: '!',
      })
    ).ok
  ).toBe(true);
  expect(() => exportMdxDocument(withPart(pkg, store.part))).toThrow(
    /Cannot map edited display text/
  );
});
test.each([
  ['weather', 'hPa'],
  ['audio', 'dB'],
])('%s supplies its own block and inline Word fallback for the same MDX', (slug, units) => {
  const options = { registry: project(slug, units) };
  const bytes = importMdxDocument(instruments, options);
  const expected = [slug + ' instrument', '12 ' + units, 'See ' + units + ' 7 below.'];
  const pkg = open(bytes);
  const components = readMdxComponents(bytes);
  expect(components.map(({ name, attributes }) => [name, attributes])).toEqual([
    ['Reading', { reading: '12' }],
    ['Stamp', { reading: '7' }],
  ]);
  expect(new Set(components.map(({ controlId }) => controlId)).size).toBe(2);
  expect(wordParagraphs(pkg).map((p) => p.text)).toEqual(expected);
  const saved = open(writeOoxmlPackage(pkg));
  expect(readMdxComponents(saved)).toEqual(components);
  const mdx = exportMdxDocument(saved);
  expect(parseDocumentSyntax(mdx)).toEqual(parseDocumentSyntax(instruments));
  expect(wordParagraphs(open(importMdxDocument(mdx, options))).map((p) => p.text)).toEqual(
    expected
  );
});

const outer = [
  '<details open>',
  '<summary>Outer **summary**</summary>',
  '',
  '<ProjectNote title="Inner" open>',
  '',
  '<details>',
  '<summary>Limits</summary>',
  '',
  'Review **load** and <lab:Mark value={readSensor()} /> before {runReport()}.',
  '',
  '</details>',
  '',
  '</ProjectNote>',
  '',
  '</details>',
].join('\n');

test('editing an atomic label updates its literal attribute and preserves preceding JSX attributes', () => {
  const source = '<Gauge title="original" {...properties} title="reading" />';
  const pkg = open(importMdxDocument(source));
  const store = new TreeDocumentStore(pkg.parts.get(pkg.mainDocumentPart)!);
  const paragraph = storyParagraphs(bodyStoryRoot(store.part)!)[0];
  expect(
    store.transact((ctx) =>
      ctx.apply({ op: 'insertText', paragraphId: paragraph.id, offset: 0, text: 'new ' })
    ).ok
  ).toBe(true);
  const saved = writeOoxmlPackage(withPart(pkg, store.part));
  const [control] = readMdxComponents(saved);
  expect(control.attributes.title).toBe('new reading');
  expect(control.jsxAttributes[0]).toMatchObject({ name: 'title', value: 'original' });
  expect(control.jsxAttributes[2]).toMatchObject({ name: 'title', value: 'new reading' });
  expect(parseDocumentSyntax(exportMdxDocument(saved))).toEqual(
    parseDocumentSyntax(source.replace('"reading"', '"new reading"'))
  );
});

test('standalone one-line disclosures fold in Word and preserve their original JSX structure', () => {
  for (const source of [
    '<details><summary>Read</summary>Some **bold** text.</details>',
    '<Fold title="Question">Answer</Fold>',
  ]) {
    const bytes = importMdxDocument(source, { disclosureTags: ['Fold'] });
    const [entry] = readMdxComponents(bytes);
    expect(entry.disclosure).toBe(true);
    expect(entry.summaryParagraphId).toBeTruthy();
    expect(parseDocumentSyntax(exportMdxDocument(bytes))).toEqual(parseDocumentSyntax(source));
  }
});

test('nested HTML and project disclosures export current Word captions and native child text', () => {
  const pkg = open(importMdxDocument(outer, { disclosureTags: ['ProjectNote'] }));
  const components = readMdxComponents(pkg);
  expect(components.map(({ name, disclosure }) => [name, disclosure])).toEqual([
    ['details', true],
    ['ProjectNote', true],
    ['details', true],
    ['lab:Mark', false],
    ['mdxTextExpression', false],
  ]);
  const paragraphs = wordParagraphs(pkg);
  const captions = components
    .filter((item) => item.disclosure)
    .map((item) => paragraphs.find((p) => p.id === item.summaryParagraphId)?.text);
  expect(captions).toEqual(['Outer summary', 'Inner', 'Limits']);
  expect(components[3]?.attributes.value).toMatchObject({
    type: 'mdxJsxAttributeValueExpression',
    value: 'readSensor()',
  });
  expect(parseDocumentSyntax(exportMdxDocument(pkg))).toEqual(parseDocumentSyntax(outer));
  const store = new TreeDocumentStore(pkg.parts.get(pkg.mainDocumentPart)!);
  const edits = [
    ['Outer summary', 'Site '],
    ['Inner', 'New '],
    ['Limits', 'Load '],
    ['Review', 'Please '],
  ];
  expect(
    store.transact((ctx) => {
      for (const [text, prefix] of edits) {
        const p = paragraphs.find((p) => p.text?.startsWith(text))!;
        ctx.apply({ op: 'insertText', paragraphId: p.id, offset: 0, text: prefix });
      }
    }).ok
  ).toBe(true);
  const expected = outer
    .replace('Outer', 'Site Outer')
    .replace('Inner', 'New Inner')
    .replace('Limits', 'Load Limits')
    .replace('Review', 'Please Review');
  const edited = withPart(pkg, store.part);
  for (const doc of [edited, writeOoxmlPackage(edited)]) {
    expect(parseDocumentSyntax(exportMdxDocument(doc))).toEqual(parseDocumentSyntax(expected));
    expect(
      readMdxComponents(doc).find(({ name }) => name === 'ProjectNote')?.attributes.title
    ).toBe('New Inner');
  }
});

test('document and attribute expressions stay inert through import, inventory, and export', () => {
  let calls = 0;
  const key = Symbol.for('mdx-document.test.expression');
  const expression = 'globalThis[Symbol.for("mdx-document.test.expression")]()';
  const source =
    'import entry from "./unloaded.mjs"\n\n' +
    '<lab:Status {...config} check={' +
    expression +
    '} />\n\n' +
    'Value {' +
    expression +
    '}.\n\n{' +
    expression +
    '}';
  Reflect.set(globalThis, key, () => ++calls);
  try {
    const bytes = importMdxDocument(source);
    const items = readMdxComponents(bytes);
    expect(items.map((item) => item.name)).toEqual([
      'mdxjsEsm',
      'lab:Status',
      'mdxTextExpression',
      'mdxFlowExpression',
    ]);
    expect(items[0]?.value).toBe('import entry from "./unloaded.mjs"');
    expect(items[1]?.jsxAttributes).toMatchObject([
      { type: 'mdxJsxExpressionAttribute', value: '...config' },
      {
        type: 'mdxJsxAttribute',
        name: 'check',
        value: { type: 'mdxJsxAttributeValueExpression', value: expression },
      },
    ]);
    expect(items.slice(2).map(({ value }) => value)).toEqual([expression, expression]);
    expect(parseDocumentSyntax(exportMdxDocument(bytes))).toEqual(parseDocumentSyntax(source));
    expect(calls).toBe(0);
  } finally {
    Reflect.deleteProperty(globalThis, key);
  }
});

for (const [placement, paragraph, start] of [
  ['block', '12 hPa', 0],
  ['inline', 'See hPa 7 below.', 4],
] as const) {
  test.each([
    { localName: 'b' },
    { localName: 'u', attributes: { val: 'single' } },
    { localName: 'color', attributes: { val: '2266AA' } },
  ])(placement + ' fallback refuses an unmappable Word $localName edit', (property) => {
    const pkg = open(importMdxDocument(instruments, { registry: project('weather', 'hPa') }));
    const id = wordParagraphs(pkg).find((p) => p.text === paragraph)!.id;
    const store = new TreeDocumentStore(pkg.parts.get(pkg.mainDocumentPart)!);
    expect(
      store.transact((ctx) =>
        ctx.apply({
          op: 'setRunProperties',
          paragraphId: id,
          start,
          end: start + 2,
          properties: [property],
        })
      ).ok
    ).toBe(true);
    const saved = writeOoxmlPackage(withPart(pkg, store.part));
    expect(wordParagraphs(open(saved)).find((p) => p.text === paragraph)).toBeDefined();
    expect(() => exportMdxDocument(saved)).toThrow(/component|fallback/i);
  });
}

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const stop of cleanups.splice(0).reverse()) stop();
});
function mount(source: string, parent = document.body) {
  const container = document.createElement('div');
  container.className = 'docx-editor';
  parent.append(container);
  const editor = createDocxEditor({ container, document: importMdxDocument(source), zoom: 1 });
  cleanups.push(() => {
    editor.destroy();
    container.remove();
  });
  return { container, editor, surface: editor.surface! };
}

test('the first layout respects authored disclosure states before the buttons mount', async () => {
  const source =
    '<details><summary>Closed</summary>Hidden data</details>\n\n<details open><summary>Open</summary>Visible data</details>';
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: importMdxDocument(source),
    configureInitialBlockView: configureMdxDisclosures,
  });
  cleanups.push(() => {
    editor.destroy();
    container.remove();
  });
  const initial = editor.surface!.collapsedContentControls();
  expect(initial).toHaveLength(1);
  expect(container.textContent).toContain('Visible data');
  expect(container.textContent).not.toContain('Hidden data');
  cleanups.push(mountMdxDisclosures(editor));
  expect(editor.surface!.collapsedContentControls()).toBe(initial);
  expect(exportMdxDocument(new Uint8Array(await editor.save()))).toContain('Hidden data');
});

test('rich summaries keep their own children and views follow only edits inside their control', async () => {
  const source = [
    '<details open>',
    '<summary>',
    '',
    'Read **first**',
    '',
    '- Label $x$',
    '',
    '</summary>',
    '',
    '<Series>',
    '',
    '| Point | Value |',
    '| --- | --- |',
    '| <Unit><Unknown>Depth</Unknown></Unit> | $y$ |',
    '',
    '</Series>',
    '',
    '</details>',
    '',
    'Tail',
  ].join('\n');
  const { container, editor, surface } = mount(source);
  const calls = new Map<string, string[]>();
  const render = ({ element, component, observeContent }: MdxComponentView) => {
    const button = element.ownerDocument.createElement('button');
    element.append(button);
    const values: string[] = [];
    calls.set(component.name, values);
    return observeContent((control) => {
      const value = contentControlTextOf(control);
      values.push(value);
      button.title = value;
    });
  };
  cleanups.push(
    mountMdxComponents(
      editor,
      new Map([
        ['Series', render],
        ['Unknown', render],
      ])
    )
  );
  cleanups.push(mountMdxDisclosures(editor));
  surface.flushPendingInput();
  const series = calls.get('Series')!;
  const unknown = calls.get('Unknown')!;
  expect(unknown).toEqual(['Depth']);
  expect(series).toHaveLength(1);
  const button = container.querySelector('[data-mdx-component="Unknown"] button');
  expect(button).not.toBeNull();
  const edit = (needle: string) => {
    const hit = editor.findMatches(needle)[0]!;
    const point = { paragraphId: hit.blockId, offset: hit.start + 1 };
    surface.setSelection({ anchor: point, head: point });
    expect(editor.exec({ type: 'insertText', text: '!' }).ok).toBe(true);
    surface.flushPendingInput();
  };
  edit('Tail');
  expect(series).toHaveLength(1);
  expect(unknown).toHaveLength(1);
  edit('Depth');
  expect(unknown).toEqual(['Depth', 'D!epth']);
  expect(series).toHaveLength(2);
  expect(series[1]).toContain('D!epth');
  expect(container.querySelector('[data-mdx-component="Unknown"] button')).toBe(button);
  const toggle = () =>
    container.querySelector<HTMLButtonElement>('button[aria-label="Read first"]')!;
  toggle().click();
  surface.flushPendingInput();
  expect(toggle().getAttribute('aria-expanded')).toBe('false');
  expect(button!.isConnected).toBe(false);
  const id = editor.findMatches('Label')[0]!.blockId;
  expect(container.querySelector(`[data-paragraph-id="${id}"]`)).not.toBeNull();
  const saved = new Uint8Array(await editor.save());
  const expected = source.replace('Depth', 'D!epth').replace('Tail', 'T!ail');
  expect(parseDocumentSyntax(exportMdxDocument(saved))).toEqual(parseDocumentSyntax(expected));
  expect(semanticDigest(open(saved).parts.values())).toEqual(
    semanticDigest(surface.session.currentPackage().parts.values())
  );
});

test('a closed disclosure keeps its leading summary table editable through save and MDX export', async () => {
  const source =
    '<details>\n<summary>\n\n| Limit | Unit |\n| --- | --- |\n| **Maximum** | <Unit>kPa</Unit> |\n\nCaption\n\n</summary>\n\nPressure reading\n\n</details>';
  const { container, editor, surface } = mount(source);
  cleanups.push(mountMdxDisclosures(editor));
  surface.flushPendingInput();
  const [details] = readMdxComponents(surface.session.currentPackage());
  expect(details!.summaryParagraphId).toBe(editor.findMatches('Limit')[0]!.blockId);
  const toggle = container.querySelector<HTMLButtonElement>('button[aria-label="Limit"]')!;
  expect(toggle.getAttribute('aria-expanded')).toBe('false');
  const painted = () =>
    Array.from(container.querySelectorAll('.docx-paragraph-fragment')).map(
      (node) => node.textContent
    );
  expect(painted()).toEqual(['Limit', 'Unit', 'Maximum', 'kPa', 'Caption']);
  const point = { paragraphId: editor.findMatches('kPa')[0]!.blockId, offset: 1 };
  surface.setSelection({ anchor: point, head: point });
  expect(editor.exec({ type: 'insertText', text: '!' }).ok).toBe(true);
  surface.flushPendingInput();
  expect(painted()).toContain('k!Pa');
  const saved = new Uint8Array(await editor.save());
  expect(parseDocumentSyntax(exportMdxDocument(saved))).toEqual(
    parseDocumentSyntax(source.replace('kPa', 'k!Pa'))
  );
  expect(semanticDigest(open(saved).parts.values())).toEqual(
    semanticDigest(surface.session.currentPackage().parts.values())
  );
  toggle.click();
  surface.flushPendingInput();
  expect(painted()).toContain('Pressure reading');
});

test('a wrapped inline namespaced component gets a separate host on each native line', () => {
  const label = 'north station '.repeat(20).trimEnd();
  const source = 'Read **the <lab:Mark title="' + label + '" /> gauge** here.';
  const { editor, surface } = mount(source);
  const views: MdxComponentView[] = [];
  const stop = mountMdxComponents(
    editor,
    new Map([
      [
        'lab:Mark',
        (view: MdxComponentView) => {
          views.push(view);
        },
      ],
    ])
  );
  cleanups.push(stop);
  surface.flushPendingInput();
  const view = views[0]!;
  const paragraph = wordParagraphs(surface.session.currentPackage()).find((item) =>
    item.text?.includes(label)
  )!;
  const start = paragraph.text!.indexOf(label);
  const layout = surface.layout();
  const control = layout.contentControls!.find(({ id }) => id === view.component.controlId)!;
  expect(control.range).toMatchObject({
    paragraphId: paragraph.id,
    start,
    end: start + label.length,
  });
  expect(control.fragments.length).toBeGreaterThan(1);
  expect(views.map((item) => item.fragmentIndex)).toEqual(
    control.fragments.map((_, index) => index)
  );
  expect(new Set(views.map((item) => item.element)).size).toBe(control.fragments.length);
  for (const item of views) {
    expect(item.pageIndex).toBe(view.pageIndex);
    expect(parseFloat(item.element.style.width)).toBeGreaterThan(0);
  }
  expect(parseDocumentSyntax(exportMdxDocument(surface.session.currentPackage()))).toEqual(
    parseDocumentSyntax(source)
  );
  stop();
  expect(views.every((item) => !item.element.isConnected)).toBe(true);
});

test.each(['disclosure', 'project'])(
  'the %s button keeps Enter for native keyboard activation',
  (kind) => {
    const { editor, container, surface } = mount(
      '<details open>\n<summary>Limits</summary>\n\n<Plot title="Native plot" />\n\n</details>'
    );
    cleanups.push(mountMdxDisclosures(editor));
    cleanups.push(
      mountMdxComponents(
        editor,
        new Map([
          [
            'Plot',
            ({ element }) => {
              const button = element.ownerDocument.createElement('button');
              button.type = 'button';
              button.textContent = 'Inspect plot';
              element.append(button);
            },
          ],
        ])
      )
    );
    surface.flushPendingInput();
    const selector =
      kind === 'disclosure' ? 'button[aria-label="Limits"]' : '[data-mdx-component="Plot"] button';
    const button = container.querySelector<HTMLButtonElement>(selector)!;
    const revision = surface.session.packageRevision();
    button.focus();
    const enter = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    button.dispatchEvent(enter);
    // The browser activates the native button only if the editor leaves its default action intact.
    expect(enter.defaultPrevented).toBe(false);
    expect(surface.session.packageRevision()).toBe(revision);
  }
);

test('trusted views keep their host and state across paint and release each hidden or destroyed host', () => {
  const source =
    '<details open>\n<summary>Instruments</summary>\n\n<Reading title="Native reading" />\n\n</details>\n\nTail';
  const { container, editor, surface } = mount(source);
  let frame: SurfacePaintFrame;
  cleanups.push(
    surface.observePaint({
      afterPaint(next) {
        frame = next;
      },
    })
  );
  const mounts: MdxComponentView[] = [];
  const released: HTMLElement[] = [];
  cleanups.push(mountMdxDisclosures(editor));
  const stop = mountMdxComponents(
    editor,
    new Map([
      [
        'Reading',
        (view: MdxComponentView) => {
          mounts.push(view);
          const button = view.element.ownerDocument.createElement('button');
          button.textContent = 'Inspect';
          view.element.append(button);
          return () => {
            released.push(view.element);
          };
        },
      ],
    ])
  );
  cleanups.push(stop);
  surface.flushPendingInput();
  expect(mounts).toHaveLength(1);
  const first = mounts[0]!;
  expect(first.editor).toBe(editor);
  expect(first.component.name).toBe('Reading');
  expect(first.element.parentElement?.closest('.docx-page')?.getAttribute('data-page-index')).toBe(
    String(first.pageIndex)
  );
  expect(first.element.contentEditable).toBe('false');
  const assertGeometry = () => {
    const box = frame.layout
      .contentControls!.find(({ id }) => id === first.component.controlId)!
      .fragments.find(({ pageIndex }) => pageIndex === first.pageIndex)!.box;
    expect(box.width).toBeGreaterThan(0);
    expect(box.height).toBeGreaterThan(0);
    for (const [prop, value] of [
      ['left', box.x],
      ['top', box.y],
      ['width', box.width],
      ['height', box.height],
    ] as const)
      expect(parseFloat(first.element.style[prop])).toBeCloseTo(value * frame.scale, 5);
  };
  assertGeometry();
  const width = parseFloat(first.element.style.width);
  first.element.firstElementChild!.setAttribute('aria-pressed', 'true');
  expect(editor.setZoom(1.5).ok).toBe(true);
  surface.flushPendingInput();
  assertGeometry();
  expect(parseFloat(first.element.style.width)).toBeGreaterThan(width);
  const tail = editor.findMatches('Tail')[0]!;
  const point = { paragraphId: tail.blockId, offset: 0 };
  surface.setSelection({ anchor: point, head: point });
  expect(editor.exec({ type: 'insertText', text: 'Changed ' }).ok).toBe(true);
  surface.flushPendingInput();
  expect(container.querySelector('[aria-pressed="true"]')?.parentElement).toBe(first.element);
  expect(mounts).toHaveLength(1);
  expect(released).toHaveLength(0);

  const toggle = container.querySelector<HTMLButtonElement>('button[aria-label="Instruments"]')!;
  toggle.click();
  surface.flushPendingInput();
  expect(released).toEqual([first.element]);
  expect(first.element.isConnected).toBe(false);
  container.querySelector<HTMLButtonElement>('button[aria-label="Instruments"]')!.click();
  surface.flushPendingInput();
  expect(mounts).toHaveLength(2);
  expect(mounts[1]!.element).not.toBe(first.element);
  editor.destroy();
  stop();
  expect(released).toEqual([first.element, mounts[1]!.element]);
});

test('scrolling mounts visible fragments, retains the pinned page, and releases offscreen fragments', async () => {
  const scroller = document.createElement('div');
  scroller.className = 'docx-editor__scroll-container';
  Object.defineProperty(scroller, 'clientHeight', { value: 115 });
  document.body.append(scroller);
  cleanups.push(() => scroller.remove());
  const { editor, surface } = mount(
    '<Chapter>\n\n' + 'Station record.\n\n'.repeat(60) + '</Chapter>',
    scroller
  );
  expect(
    editor.exec({ type: 'setPageSetup', pageHeight: 2000, marginTop: 140, marginBottom: 140 }).ok
  ).toBe(true);
  surface.flushPendingInput();
  const pages = surface.layout().pages;
  expect(pages.length).toBeGreaterThan(4);
  const last = pages.at(-1)!;
  const mounted = new Map<number, HTMLElement>();
  const stop = mountMdxComponents(
    editor,
    new Map([
      [
        'Chapter',
        ({ element, pageIndex }) => {
          mounted.set(pageIndex, element);
          return () => {
            mounted.delete(pageIndex);
          };
        },
      ],
    ])
  );
  cleanups.push(stop);
  const pinned = mounted.get(0)!;
  expect(pinned.isConnected).toBe(true);
  const lastIndex = pages.length - 1;
  expect(mounted.has(lastIndex)).toBe(false);
  async function scroll(top: number, ready: () => boolean) {
    scroller.scrollTop = top;
    scroller.dispatchEvent(new Event('scroll'));
    const deadline = performance.now() + 2000;
    while (!ready() && performance.now() < deadline)
      await new Promise((resolve) => setTimeout(resolve, 10));
    expect(ready()).toBe(true);
  }
  await scroll((last.box.y * 96) / 72, () => mounted.has(lastIndex));
  const far = mounted.get(lastIndex)!;
  expect(far.isConnected).toBe(true);
  expect(mounted.get(0)).toBe(pinned);
  await scroll(0, () => !mounted.has(lastIndex));
  expect(far.isConnected).toBe(false);
  expect(mounted.get(0)).toBe(pinned);
  stop();
  expect(mounted.size).toBe(0);
});

test('disclosure buttons name their native caption and local folds preserve the saved document and history', async () => {
  const { editor, container, surface } = mount(
    '<details>\n<summary>Limits</summary>\n\nHidden body\n\n</details>\n\nTail'
  );
  const digest = async () =>
    semanticDigest(open(new Uint8Array(await editor.save())).parts.values());
  const before = await digest();
  const rev = surface.session.packageRevision();
  const stop = mountMdxDisclosures(editor);
  cleanups.push(stop);
  const button = () => container.querySelector<HTMLButtonElement>('button[aria-label="Limits"]')!;
  const body = editor.findMatches('Hidden body')[0]!;
  const visible = () => container.querySelector(`[data-paragraph-id="${body.blockId}"]`) !== null;
  surface.flushPendingInput();
  expect(button().type).toBe('button');
  expect(button().tabIndex).toBe(0);
  expect(button().getAttribute('aria-expanded')).toBe('false');
  expect(visible()).toBe(false);
  button().click();
  surface.flushPendingInput();
  expect(button().getAttribute('aria-expanded')).toBe('true');
  expect(visible()).toBe(true);
  button().focus();
  button().click();
  surface.flushPendingInput();
  expect(container.ownerDocument.activeElement === button()).toBe(true);
  expect(await digest()).toEqual(before);
  expect(surface.session.packageRevision()).toBe(rev);
  expect(editor.snapshot().canUndo).toBe(false);
  stop();
  surface.flushPendingInput();
  expect(visible()).toBe(true);
  expect(button()).toBeNull();
});

async function checkDisclosureViewport(wrapperOverflow: string) {
  const viewport = document.createElement('section');
  const wrapper = document.createElement('div');
  viewport.style.overflowY = wrapperOverflow === 'document' ? 'visible' : 'auto';
  wrapper.style.overflowY = wrapperOverflow === 'document' ? 'clip' : wrapperOverflow;
  document.body.append(viewport);
  viewport.append(wrapper);
  cleanups.push(() => viewport.remove());
  const { editor, container, surface } = mount(
    '<details>\n<summary>Observation</summary>\n\nFirst reading\n\nSecond reading\n\n</details>\n\nFollowing text',
    wrapper
  );
  const doc = container.ownerDocument;
  const win = doc.defaultView!;
  const restore: (() => void)[] = [];
  function replace(target: object, key: string, value: unknown) {
    const prior = Object.getOwnPropertyDescriptor(target, key);
    Object.defineProperty(target, key, { value, configurable: true });
    restore.push(() => {
      if (prior) Object.defineProperty(target, key, prior);
      else Reflect.deleteProperty(target, key);
    });
  }
  const captures: { update: () => void; finish: () => void; ready: () => void }[] = [];
  let reversals = 0;
  replace(viewport, 'clientHeight', 360);
  replace(viewport, 'scrollHeight', 1800);
  replace(wrapper, 'clientHeight', 1800);
  replace(doc, 'scrollingElement', doc.documentElement);
  const scrollNode = wrapperOverflow === 'document' ? doc.documentElement : viewport;
  const scrollTarget = wrapperOverflow === 'document' ? doc : viewport;
  const scale = wrapperOverflow === 'clip' ? 0.65 : 1;
  replace(win, 'matchMedia', () => ({
    matches: true,
    addEventListener() {},
    removeEventListener() {},
  }));
  replace(doc, 'startViewTransition', (update: () => void) => {
    const ready = Promise.withResolvers<void>();
    const finished = Promise.withResolvers<void>();
    captures.push({ update, ready: ready.resolve, finish: finished.resolve });
    return { ready: ready.promise, finished: finished.promise, skipTransition: finished.resolve };
  });
  replace(doc, 'getAnimations', () => [
    { animationName: 'docx-mdx-move-following', playState: 'running', reverse: () => reversals++ },
  ]);
  const rect = win.HTMLElement.prototype.getBoundingClientRect;
  replace(win.HTMLElement.prototype, 'getBoundingClientRect', function (this: HTMLElement) {
    if (
      this === viewport ||
      this.classList.contains('docx-pages') ||
      this.classList.contains('docx-page')
    ) {
      const page = surface.publishedLayout().pages[0]!;
      return new win.DOMRect(0, 0, page.box.width, page.box.height * scale);
    }
    if (this.hasAttribute('data-docx-mdx-disclosure'))
      return new win.DOMRect(
        0,
        (this.getAttribute('aria-expanded') === 'true' ? 12 : 0) - scrollNode.scrollTop,
        22,
        22
      );
    return rect.call(this);
  });
  try {
    cleanups.push(mountMdxDisclosures(editor));
    surface.flushPendingInput();
    const button = container.querySelector<HTMLButtonElement>('button[aria-label="Observation"]')!;
    const saved = exportMdxDocument(surface.session.currentPackage());
    const layout = surface.publishedLayout();
    const closed = layout.contentControls
      ?.find((record) => record.id === surface.collapsedContentControls()[0]?.controlId)!
      .fragments.at(-1)!;
    const sheet = layout.pages[closed.pageIndex]!;
    const closedPercent =
      (100 * (sheet.contentBox.y - sheet.box.y + closed.box.y + closed.box.height)) /
      sheet.box.height;
    button.focus();
    button.click();
    expect(captures).toHaveLength(1);
    const snapshotPane =
      wrapperOverflow === 'document'
        ? container.querySelector<HTMLElement>('.docx-pages')!
        : viewport;
    expect(snapshotPane.style.viewTransitionName).toBe('docx-mdx-paper');
    expect(doc.documentElement.classList.contains('docx-mdx-opening')).toBe(true);
    captures[0]!.update();
    const motionEdge = doc.documentElement.style.getPropertyValue('--docx-mdx-closed-edge');
    expect(motionEdge.endsWith('%')).toBe(true);
    expect(Number.parseFloat(motionEdge)).toBeCloseTo(closedPercent, 5);
    expect(scrollNode.scrollTop).toBe(12);
    scrollTarget.dispatchEvent(new win.Event('scroll'));
    expect(doc.documentElement.classList.contains('docx-mdx-opening')).toBe(true);
    captures[0]!.ready();
    await Promise.resolve();
    expect(button.getAttribute('aria-expanded')).toBe('true');
    expect(surface.collapsedContentControls()).toHaveLength(0);
    button.click();
    expect(reversals).toBe(1);
    expect(captures).toHaveLength(1);
    expect(button.getAttribute('aria-expanded')).toBe('false');
    expect(surface.collapsedContentControls()).toHaveLength(0);
    captures[0]!.finish();
    await Promise.resolve();
    expect(surface.collapsedContentControls()).toHaveLength(1);
    expect(doc.activeElement).toBe(button);
    expect(doc.documentElement.classList.contains('docx-mdx-opening')).toBe(false);
    expect(exportMdxDocument(new Uint8Array(await editor.save()))).toBe(saved);
    button.click();
    captures[1]!.update();
    scrollNode.scrollTop += 32;
    scrollTarget.dispatchEvent(new win.Event('scroll'));
    expect(doc.documentElement.classList.contains('docx-mdx-opening')).toBe(false);
    expect(button.getAttribute('aria-expanded')).toBe('true');
  } finally {
    scrollNode.scrollTop = 0;
    for (const undo of restore.reverse()) undo();
  }
}

test.each(['clip', 'hidden', 'visible', 'document'])(
  'rapid input reverses the disclosure inside its scroll viewport (%s)',
  checkDisclosureViewport
);
