# React document blocks proof

This experimental example embeds application React components in the canonical Word document flow. It uses the Apache core and React adapter. It registers its own review module with Apache core exports.

The fixture contains a phase, a nested activity, native editable paragraphs and a table, and an interactive quiz. The quiz includes SVG, a select control, and expandable content with an independent textarea.

## Run

From the repository root:

```bash
bun install
bun run --filter @docx-editor.dev/fonts build
bun run --filter docx-editor-example-react-blocks dev
```

Open `http://127.0.0.1:5188`.

Edit the paragraphs and table cells. Change the activity duration, select a prediction, and use **Save and reopen DOCX**. Use **Propose AI text edit** to create a deterministic native tracked replacement inside the activity. No model call is required for this test.

Click the activity's local-state button, then use **Push blocks to next page**. The same React root moves with the block and retains its local state.

The demo loads packaged Calibri-compatible font files for measurement and painting. It does not depend on the system Calibri fallback.

## Document and render contracts

- A block is a nested rich-text Word content control. Its tag selects a trusted application renderer. Document content cannot supply executable components.
- Word paragraphs and tables remain canonical and editable. React owns their visual wrappers, through a native-content slot.
- `surface.observePaint()` restores borrowed elements before the core paints and reattaches them before the core restores the DOM selection.
- React roots use the canonical block identity and layout fragment ordinal. Page movement retains roots. Virtualized fragments are parked.
- Native UI event isolation keeps form-control input out of the document's input handler. Native editable content continues through the normal input path.
- `ResizeObserver` measures the React header. The adapter records its height as exact Word paragraph spacing, so the canonical layout reserves the space.
- Word saves ordinary fallback text and nested content. React component code and browser DOM are not serialized. This fixture stores small typed payloads in the control alias; a production schema should use custom XML for larger data.

The registry accepts React components with `BlockViewProps`. Components can compose arbitrary browser UI around the native-content slot. A container can span pages through fragments. An opaque widget must fit a page; automatic fragmentation of arbitrary React DOM is outside this proof.

Each fragment uses an independent React root. Registered renderers must supply the application providers they require. Unregistered block tags retain their native Word fallback.

## Verify

```bash
bun run --filter @docx-editor.dev/i18n build
bun run --filter docx-editor-example-react-blocks typecheck
bun test packages/core/src/editor/__tests__/surface-paint-observer.test.ts
bunx playwright test --config e2e/react-blocks.config.ts
```

To use an installed Chromium executable, set `PLAYWRIGHT_CHROMIUM_EXECUTABLE=/usr/bin/chromium` for the browser test command.

The browser tests exercise real typing, native tracked changes, accept/reject, DOCX save/reopen, measured widget resizing, and React state across page movement.

The optional collaboration compatibility test creates two real Yjs replicas. It covers concurrent edits, reconnect, undo/redo, remote selections, deletion while another participant edits, and matching saved semantic digests:

```bash
bun test examples/react-blocks/test/collaboration.test.ts
```

That test evaluates the upstream collaboration runtime under the EigenPal Pro License. It is a development dependency. The runnable demo does not import it.

## Scope limits

This example does not establish full Plate parity. Production work includes block insertion and clipboard semantics, metadata storage, property suggestions, AI multi-wave diff reconciliation, large-document profiling, IME and touch interaction, error isolation, and desktop Word interoperability testing.

The editor API snapshot and a changeset cover the generic core paint-observer extension. No production integration in Sciobot is changed.
