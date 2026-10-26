# Native source citation chips

X11 uses the existing native `source` payload codec and literal shared source aliases. The application provenance renderer starts in `7b4b37d30` and `287f6fedf`. Book chips open the existing provenance modal with book ID, title, author, and URL. Web chips expose HTTP(S) links. Opening provenance leaves the native document untouched.

## Native inline review

`payload-fallback.ts` resolves the native control and its containing paragraph. It derives the current offset from native identity; it never searches by label. The text planner receives only the fallback text and its mapped offset. Tracked insertions explicitly name the owning content control. Multiple labels in one paragraph execute from right to left.

Staging, decisions, external review, and content proofs use the same control boundary. Unrelated prose and adjacent citations retain independent native revisions and history. Native locks apply to the source control, its ancestors, and generated label owners. The block and generated paragraph paths retain their existing behavior.

The focused native command is:

```bash
bun test --preload ./scripts/test/preload-dom.ts \
  packages/sciobot/src/review/recompose-inline-sources.test.ts \
  packages/sciobot/src/review/recompose-payloads.test.ts \
  packages/sciobot/src/review/track-atomic.test.ts \
  packages/sciobot/src/review/reconcile-external.test.ts \
  packages/sciobot/src/review/recompose-payloads-collaboration.test.ts
```

All 79 cases and 1,094 assertions pass in `/tmp/x11-native-final-batch.log`. Fourteen citation cases include both decisions, two waves, identical neighboring labels, complete opaque payloads, same-paragraph peer text, and independent source groups. Eight shared cases cover both client orders and reconnect receivers, both actors' histories, actual remote selection identity, exact caret geometry, saved convergence, and cold rooms. Metadata-only citation changes retain their original native text.

Private types pass 2,434 files. Scoped lint and formatting pass. The compatibility decision is `.collaboration/changes/sciobot-inline-citation-fallback.json`. It retains existing operations, atomic records, and the collaboration format. This private adapter change adds no public API or publishable package release.

The previous pinned adapter reads exact original and proposed views from the repaired writer:

```bash
bun --preload ./scripts/test/preload-dom.ts \
  scripts/sciobot-port/verify-inline-citation-reader.ts \
  /home/lofcz/.cache/sciobot-docx-native/port-current-integration/packages/sciobot/dist/index.js
```

The previous adapter is `4f0029f0b8df49f1aa966c43ce52a7a97b8443b0`. The verified reader log is `/tmp/x11-earlier-reader-final.log`. Public core runtime bytes remain the integration baseline until the installed check below.

## Application checks

The initial UI checkpoint passes six browser cases and 66 assertions. It covers provenance, offline deletion, independent history, geometry, and reopening. Its log is `/tmp/x11-browser-ui-final.log`. Fifteen units pass; frontend and harness types pass 6,851 and 2,426 files.

Application `e9eeaa81f` installs source runtime `926527dd155f7bd647d538d7ecbcadea2683d4ec` from the recorded build. The installed private archive is `vendor/docx-native/sciobot-docx-c7b05ecf1598.tgz`. All five archive hashes pass. The four public archives retain exact bytes. The frozen install and single physical core resolution pass. The private build validates 335 source modules with native core external. The shared format stays `docx-collaboration:1.9.15.1`; private field protocol stays 3.

The completed UI/offline command verifies six installed cases:

```bash
DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/x11-native-integration \
DOCX_HARNESS_PORT=8122 DOCX_HARNESS_URL=http://127.0.0.1:8122 \
bun test --timeout 600000 scripts/docx-native/harness/source-citations.test.ts
```

Those six cases pass in `/tmp/x11-installed-review-verified.log`. That combined invocation later exposes stale Original-view provenance; it is not counted as a passing full invocation. The final review correction passes this focused command:

```bash
DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/x11-native-integration \
DOCX_HARNESS_PORT=8122 DOCX_HARNESS_URL=http://127.0.0.1:8122 \
bun test --timeout 600000 --test-name-pattern 'citation provenance, AI replacement' \
  scripts/docx-native/harness/source-citations.test.ts
```

All four cases and 112 assertions pass in `/tmp/x11-installed-ai-final.log`. Together these commands verify ten distinct browser cases. Both client orders and both decisions preserve complete source payloads and opaque data. Two real tool waves withdraw earlier groups and retain peer typing through both histories. Accepted citations retain URL, title, author, media type, numeric book ID, and trust. Rejected citations retain their complete originals. Cold DOCX reopening retains both results.

Mixed Word revisions retain native text and attribution. Transparent provenance hit targets cover each published native control fragment. Tests compare all four rectangle coordinates against native control chrome and retain native revision spans. Original view opens original provenance. All Markup uses the proposed citation. Citation views disable React Compiler memoization for imperative native view changes; changing the review view does not change the canonical package revision. The original assertions and deadlines remain intact.

Final frontend/harness types pass 6,852/2,427 files. Fifteen application units pass. Scoped lint passes. Logs are `/tmp/x11-app-final-fe.log`, `/tmp/x11-review-ink-harness.log`, `/tmp/x11-installed-units.log`, and `/tmp/x11-app-final-lint.log`. The former AI reproducer lives in the standard citation suite. General shared MDX editing remains X34; legacy source/book citations use the native source codec.
