# Native callout and story renderers

Runtime: `f5bb40c9c703e61abef8cec05021d02eda715f78`. Application: `fad7d96a2232a048418c4d3069212bc27c1fda28`. Branch: `codex/react-block-proof` in both repositories.

## Implemented behavior

The application registers callouts and stories as native-content renderers. A shared frame paints each fragment behind its native slot. Labels and controls occupy the page margin. Native text remains visible and clickable. A browser screenshot confirms the nested frame presentation.

Callouts reuse the application's appearance helpers and emoji picker. The appearance helper now accepts plain data, without Plate types. Variants, icons, explicit background colors, and opaque metadata stay canonical. Emoji changes use the existing gated payload writer in edit mode. Other modes disable that control. Stories reuse the localized label and book icon. Their metadata and body remain separate from host chrome.

The real Markdown import exposed a missing authoring mapping. Registered tags were readable, but target planning could not create these payload types. The private planner now recognizes callout, story, and solution containers. It preserves native owners during later text targets. Callout `type` and legacy `variant` attributes follow the application normalization and precedence rules. This adds no new shared encoding or public automation API.

## Verification

From the fork:

```sh
bun test --timeout 60000 --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/container-authoring.test.ts packages/sciobot/src/diff/domain-containers.test.ts packages/sciobot/src/projection/container-tags.test.ts
```

All 47 cases pass, with 214 assertions. The log is `/tmp/x13-native-authoring-final.log`. The new cases check creation, nested lists, retained identities in a second wave, saved reopening, metadata, and variant aliases. Private types pass across 2,428 files. All normal commit hooks pass. Compatibility is recorded in `.collaboration/changes/sciobot-native-container-authoring.json`. The private Changeset is empty because the SDK is not published.

From the isolated application worktree:

```sh
DOCX_NATIVE_SNAPSHOT=/home/lofcz/.cache/sciobot-docx-native/port-current-integration bun test --timeout 600000 scripts/docx-native/harness/containers.test.ts scripts/docx-native/harness/retiming.test.ts
bun check -p tsconfig.json
bun check -p scripts/docx-native/tsconfig.json
bun install --frozen-lockfile --ignore-scripts
```

All eight browser cases pass, with 960 assertions. The log is `/tmp/x13-containers-browser-final.log`. Three cases exercise the new containers. Five retain retiming coverage against the updated native ownership profile. Frontend types pass across 6,838 files. Harness types pass across 2,411 files. Scoped formatting and lint pass. The five archives match `SOURCE.json` and the package manifest. App, SDK, and converter resolve one core runtime. The private archive is `sciobot-docx-268869a2513e.tgz`.

The fixture imports a story with a nested callout and a two-level native list. Real pointer and keyboard actions edit the story, callout, and nested list. Both client orders cover disconnected edits, independent undo/redo, deletion during peer typing, unchanged payloads and owner identities, exact caret geometry, saved native convergence, and reopening. Assertions retain original metadata and Markdown nesting. Initial failed setup identified the missing authoring mapping; the passing run uses the installed immutable fix.

The evaluated collaboration transport remains test-only under the EigenPal Pro License. The production components import Apache core and the private SDK. Solution reveal controls remain X15. R19 retains its separate remaining ownership and application checks.
