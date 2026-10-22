# Contracts and reference fixtures

Freeze observable Plate behavior before replacing its implementation. These leaves provide independent expected results.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

## C01

- [x] Capture the seven AI parity fixtures

**Depends:** None.

**Primary files:** `fork:packages/sciobot/src/__tests__/fixtures/ai-pairs.json`

**Read:** [S03](sources.md#s03), [S04](sources.md#s04)

**Change:** Capture the seven before/after pairs from ai-review-diff.test.ts. Include normalized Plate results and original trees without generated IDs.

**Close when:** A fixture validator finds seven named pairs. Record source commits and content hashes, including dirty source files.

**Check:** Validate fixture schema, source coverage, and independently captured expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `bun scripts/sciobot-port/capture-ai-reference.mjs` executes the application Plate implementation and captures seven cases. Direct output, accepted review, and atomic undo/redo agree. The fixture records both source revisions and SHA-256 hashes, including installed packages and dirty application files. `reference-fixtures.test.ts`: two tests, 97 assertions pass. Scoped Prettier and oxlint pass.

## C02

- [x] Capture document tool request and result contracts

**Depends:** None.

**Primary files:** `fork:packages/sciobot/src/__tests__/fixtures/tool-contracts.json`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05)

**Change:** Record read, search, edit, and set examples. Include defaults, errors, result echoes, path routing, and the is_suggesion wire spelling.

**Close when:** Each tool has successful and failing examples. Include direct-default behavior and explicit suggestion behavior.

**Check:** Validate fixture schema, source coverage, and independently captured expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `bun scripts/sciobot-port/capture-tool-reference.mjs` captures 12 schema examples, eight read/search examples, and 11 source-executed executor examples. Includes exact echoes, explicit suggestion mode, syntax/schema errors, path enum validation, and declared/effective read defaults. Scope excludes mounted hook routing, assigned to I06. Fixture tests pass five cases and 210 assertions. Typecheck, scoped Prettier, and oxlint pass. Source hashes are in `tool-contracts.json`.

## C03

- [x] Create the private application adapter package

**Depends:** None.

**Primary files:** `fork:packages/sciobot/{package.json,tsconfig.json,src/index.ts}`; `fork:bun.lock`

**Read:** [S25](sources.md#s25)

**Change:** Create private @sciobot/docx with one core peer dependency. Add scoped typecheck and test commands without public document APIs.

**Close when:** The empty package typechecks and imports the existing core. Package resolution finds one core instance.

**Check:** Run scoped build and typecheck checks. Save the inspected dependency or bundle evidence.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `bun run --filter @sciobot/docx build`, `typecheck`, and `test` pass. `package-boundary.test.ts` passes five assertions for the private package and one core identity. Scoped Prettier and oxlint pass. This closes package setup only.

## C04

- [x] Port the pure exact replacement matcher

**Depends:** [C03](01-contracts.md#c03), [C05](01-contracts.md#c05)

**Primary files:** `fork:packages/sciobot/src/tools/find-replace.ts`; `sibling find-replace.test.ts`

**Read:** [S06](sources.md#s06)

**Change:** Port exact, whitespace, and typography matching. Preserve uniqueness, replacement counts, literal dollar sequences, and atomic failure.

**Close when:** Every C05 expected output matches. Ambiguous and count-mismatched patches leave the input unchanged.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. The private matcher ports the application shared engine and pins diff 9.0.0. All 20 captured calls match full reference outputs; ambiguity and count mismatch return no partial content. Package tests pass 25 cases and 209 assertions. Package build, typecheck, scoped Prettier, and oxlint pass. Fixture: `find-replace.json`.

## C05

- [x] Capture replacement matcher regression cases

**Depends:** None.

**Primary files:** `fork:packages/sciobot/src/__tests__/fixtures/find-replace.json`

**Read:** [S06](sources.md#s06)

**Change:** Extract the existing matcher regressions into framework-independent input/output records. Include Czech quotes, entities, indentation, and deletion newlines.

**Close when:** Every existing matcher test maps to one fixture record with an explicit expected string or error.

**Check:** Validate fixture schema, source coverage, and independently captured expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `bun scripts/sciobot-port/capture-matcher-reference.mjs` executes all 17 original regression tests and captures all 20 calls, including full success strings and errors. The fixture records source hashes and the application revision. Fixture validation passes 181 assertions across three tests. Scoped Prettier and oxlint pass.

## C06

- [x] Expand the production feature coverage ledger

**Depends:** None.

**Primary files:** `fork:docs/engineering/sciobot-port/coverage.md`; `evidence/production-coverage.json`; `fork:scripts/sciobot-port/check-production-coverage.mjs`

**Read:** [S01](sources.md#s01), [S17](sources.md#s17), [S21](sources.md#s21), [S22](sources.md#s22), [S23](sources.md#s23), [S27](sources.md#s27)

**Change:** Trace each kit to production commands and saved node types. Add missing leaves for uncovered actions or attributes.

**Close when:** Every imported workspace kit and shared renderer has a source, task IDs, and a pending verification cell.

**Check:** Run `bun scripts/sciobot-port/check-production-coverage.mjs` and the backlog validator. Check rejection of missing members, missing sources, missing print bindings, and changed source hashes.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78` and application reference `da9081a`, including current dirty source files. The [production inventory](evidence/production-coverage.json) traces all 37 workspace members, 26 static kits and their bound renderers, and 13 print overrides to definitions, production routes, saved content, task IDs, and pending production verification. The standalone audit verifies 196 source hashes without loading the application. Four negative controls reject a missing workspace member, missing source file, missing print override, and changed source hash. The main backlog validator runs this audit too. P20, U15, and U16 now cover the previously unmapped chat quote workflow. U01 reflects the actual selection toolbar; U03 preserves disabled context menus; U02 retains locale and capability gates. Existing static comment support stays mapped without adding a live comment UI. This closes source coverage only. Feature parity and cutover remain open. Scoped Prettier, oxlint, both validators, and diff checks pass. No document, API, shared format, or production runtime changes occur.

## C07

- [x] Define revision and edit result types

**Depends:** [C02](01-contracts.md#c02), [C03](01-contracts.md#c03)

**Primary files:** `fork:packages/sciobot/src/contracts/{document,executor}.ts`; `sibling contract tests`

**Read:** [S02](sources.md#s02), [S05](sources.md#s05), [S25](sources.md#s25)

**Change:** Define private FileExecutor-compatible types, document identity, revision tokens, edit modes, and typed failures. Keep models free of live editor objects.

**Close when:** C02 examples typecheck against these types. A stale revision and unsupported edit have distinct serializable failures.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. Private contracts preserve all four wire shapes, per-path FileExecutor results, and the is_suggesion spelling. Six captured successful schema requests and all 11 executor results satisfy compile-time types. Three contract tests pass ten assertions, including distinct JSON-safe stale and unsupported errors. Package typecheck, scoped Prettier, and oxlint pass. Revision identity includes document, load epoch, and native revision.

## C08

- [x] Port domain identity and property diff rules

**Depends:** [C03](01-contracts.md#c03)

**Primary files:** `fork:packages/sciobot/src/diff/domain-strategies.ts`; `sibling domain-strategies.test.ts`

**Read:** [S09](sources.md#s09), [S10](sources.md#s10)

**Change:** Port container identity, activity name/duration update properties, and atomic node rules. Ignore generated IDs and listStart.

**Close when:** Changing activity duration preserves its identity. Changing phaseType changes identity. Missing old properties remain distinguishable from undefined.

**Check:** Add or extend a focused Bun test with explicit expected results.

**Claim:** codex, branch `codex/react-block-proof`. **Evidence:** Native baseline `9770e78`, local working tree. `domain-strategies.test.ts` passes six tests and 24 assertions for activity updates, phase/title/curriculum/source identity, absent old properties, atomic blocks, ignored IDs/list starts, and prototype-name fallback. The implementation imports no Plate runtime. Package typecheck, scoped Prettier, and oxlint pass.
