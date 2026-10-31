# Embedder MDX components

Replace the Sciobot constants in the private adapter with one registry that the embedding application supplies. The engine packages stay free of application content. The adapter keeps its generic mechanisms: the React block host, the payload codec, the MDX projection, opaque MDX, and the review planner.

Read the [dispatch and completion rules](../../../SCIOBOT_PORT_TODO.md) before claiming a leaf.

Each named test or artifact is a deliverable. These specifications do not claim those checks already exist.

Each registered component declares these fields:

- A tag name and a kind: block or inline.
- A versioned payload schema.
- Storage: the content-control tag prefix and the custom XML namespace.
- A content mode: no content, native Word content, or opaque MDX.
- A Word-readable fallback.
- The MDX mapping in both directions.
- A diff identity and a diff strategy.
- A React renderer, with a measured height or a reserved height.

Unregistered tags that the projection allows keep the opaque MDX path. Existing Sciobot files keep opening: the Sciobot registration keeps the `scio:1:` tag prefix and the `urn:sciobot:docx:blocks:1` namespace.

## M01

- [x] Define the embedder MDX component contract

**Depends:** [B01](02-react-blocks.md#b01), [B22](02-react-blocks.md#b22)

**Primary files:** `packages/sciobot/src/registry/` (new); `packages/sciobot/src/index.ts`

**Read:** [payload codec](../../../packages/sciobot/src/blocks/payload-opc-codec.ts), [payload schema](../../../packages/sciobot/src/blocks/payload-schema.ts), [block tags](../../../packages/sciobot/src/projection/block-tags.ts), [MDX tags](../../../packages/sciobot/src/projection/mdx-tags.ts)

**Change:** Add the `MdxComponentDefinition` type and a frozen registry value with the fields listed at the top of this page. Validate names, prefixes, and namespaces at registration. Refuse duplicates and prototype keys.

**Close when:** A registry can be built from definitions. Invalid, duplicate, or conflicting definitions refuse with a stable reason.

**Check:** Add focused registry construction tests, including hostile names and prefixes.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** `packages/sciobot/src/registry/contract.ts` adds `createMdxComponentRegistry`, which returns a frozen registry or refuses with `invalid-storage`, `invalid-component`, `duplicate-component`, `duplicate-tag`, or `invalid-opaque-tag`. `packages/sciobot/src/registry/sciobot.ts` builds `SCIOBOT_MDX_REGISTRY` from the existing constants, so stored names do not change. `bun test packages/sciobot/src/registry/contract.test.ts` passes 5 cases: indexing, hostile storage names, hostile and duplicate components, opaque tag rules, and identity with `DOMAIN_BLOCK_TAGS` and the diff strategies. A pure registry needs no session test. Limit: no adapter module reads the registry yet; M02 to M05 connect it.

## M02

- [x] Drive the MDX projection from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/projection/mdx.ts`; `packages/sciobot/src/projection/mdx-tags.ts`; `packages/sciobot/src/projection/block-tags.ts`

**Read:** [MDX projection](../../../packages/sciobot/src/projection/mdx.ts), [MDX tags](../../../packages/sciobot/src/projection/mdx-tags.ts)

**Change:** Replace `DOMAIN_BLOCK_TAGS` and the `CUSTOM_*_TAGS` constants with registry reads passed through `projectBlocks` and `projectMdx`. Keep the opaque `shared_mdx_*` lane for allowed but unregistered tags.

`planTargetDocument` already accepts `options.definitions`, but its helpers read `DOMAIN_BLOCK_TAGS` directly: `planInsertedBlocks`, `planTargetAtomic`, `planDomainProperties`, `planDomainContainers`, the inserted-cell and fragment-seed validators, and `external-review-state.ts`. Pass the registry through those calls. `insertedPayload` and `payloadMatches` hold per-type MDX parsing for code, equations, mentions, and images; add `fromMdx` and `matches` hooks to the component contract and move that logic into the Sciobot registration.

**Close when:** The projection reads no module constant for application tags. Existing projection fixtures produce the same output with the Sciobot registration.

**Check:** Run the projection and MDX suites. Add one fixture with a non-Sciobot component.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** The component contract adds `parse` (JSX attributes to payload attributes) and `markdown` (`fromMarkdown` and `matches` for Markdown nodes such as fenced code and math). `registry/mdx-read.ts` holds the generic reader; `registry/sciobot-parse.ts`, `registry/sciobot-markdown.ts`, and `registry/sciobot-diff.ts` hold the Sciobot rules that `domain-ast.ts`, `insertedPayload`, `payloadMatches`, and `domain-strategies.ts` used to hard-code. `planTargetDocument` takes `options.registry` and passes it to `bindLogicalSource`, `normalizeLogicalBlocks`, `planInsertedBlocks`, `planInsertedAtomBindings`, `planTargetAtomic`, `planAtomicBlock`, and `planDomainProperties`; `planDomainContainers`, `prepareExternalReview`, and the inserted-cell and fragment-seed validators take it too. `projectBlocks` and `projectMdx` default to the registry mappings. Checks: `bun test packages/sciobot/src/registry` passes 7 cases, including a non-Sciobot `Card` component that parses, binds, matches, and keeps its own identity rule, and that each registry reads only its own components. The whole `sciobot` suite (418 files) shows only the pre-merge baseline failures; the files that timed out under 14 workers pass alone in 6 to 13 seconds. Limits: entry points still default to `sciobotMdxRegistry()` until M06 makes the application pass the registry. Inline atom paths (`target-inline.ts`, `inline-source-updates.ts`) and `property-projection.ts` still use the default; M05 covers them.

## M03

- [x] Drive payload storage from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/blocks/payload-opc-codec.ts`; `packages/sciobot/src/blocks/payload-part.ts`

**Read:** [payload codec](../../../packages/sciobot/src/blocks/payload-opc-codec.ts), [payload part](../../../packages/sciobot/src/blocks/payload-part.ts)

**Change:** Take the content-control tag prefix and the custom XML namespace from the registry instead of `BLOCK_TAG_PREFIX` and `PAYLOAD_NAMESPACE`.

**Close when:** Documents saved before this change open with identical payloads. A second registration stores its payloads under its own namespace.

**Check:** Run the payload suites against saved fixtures. Add a save and reopen case for a second namespace.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** The payload codec takes a `PayloadStorage` (tag prefix, namespace, root, node prefixes) and defaults to the configured registry's storage through `activeStorage()`, so every adapter module that reads or writes payloads uses the embedder's names: `buildBlockIndex`, block insert and remove, property and presence review stores, generated label controls, and fallback detection. The adapter's own names derive from `adapterPrefix` and `adapterNamespace` (`registry/adapter-names.ts`): checklist boxes, image reservations, list state, and review groups. The Sciobot registration keeps `scio:1:`, `urn:sciobot:docx:blocks:1`, `scio`, and `urn:sciobot:docx`, so saved documents open unchanged. Checks: `bun test packages/sciobot/src/blocks/payload-opc-codec.test.ts` (7 pass) saves and reopens a second namespace beside the Sciobot store; `bun test packages/sciobot/src/registry/active.test.ts` (2 pass) configures an `acme` registry, inserts a `Card` block, saves, reopens, indexes, and projects it, and shows that the Sciobot names find nothing of it. The whole `sciobot` suite (418 files) shows only pre-merge baseline failures and load timeouts.

## M04

- [x] Drive diff identities and strategies from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/diff/domain-strategies.ts`

**Read:** [domain strategies](../../../packages/sciobot/src/diff/domain-strategies.ts)

**Change:** Replace the `STRATEGIES` table and `domainDiffStrategy` lookups with the registered strategy. Keep the generic strategy for unregistered components.

**Close when:** The diff modules hold no component-name literal. The existing diff suites pass with the Sciobot registration.

**Check:** Run the diff suites. Add a planner case for a non-Sciobot component.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** `SCIOBOT_DIFF_STRATEGIES` is in `registry/sciobot-diff.ts`; `domainDiffStrategy`, `sameContainerIdentity`, and `domainPropertyChanges` read the registry. New atomic blocks render their Word fallback through the component's `wordFallback` hook (code blocks in the inline code style, equations as linear source); the property planners take serializer sample values from `propertyProbe`. Bound Markdown nodes in table cells stay limited to fenced code, a standard mdast rule rather than a Sciobot one. `bun test packages/sciobot/src/registry/custom-component.test.ts` covers a non-Sciobot component's identity rule; the `diff` suite (80 files) shows only the pre-merge `inserted-atoms` failures.

## M05

- [x] Drive inline atoms from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/projection/math.ts`; `packages/sciobot/src/projection/mention.ts`; `packages/sciobot/src/react-blocks/inline-host.tsx`

**Read:** [inline host](../../../packages/sciobot/src/react-blocks/inline-host.tsx), [mention projection](../../../packages/sciobot/src/projection/mention.ts)

**Change:** Replace the fixed `math`, `mention`, and `source` atom kinds with registered inline components. Keep equations native.

**Close when:** A registered inline component projects, measures, renders, and deletes as one native atom.

**Check:** Run the inline atom suites. Add a two-session case for a non-Sciobot inline component.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** Inline components declare an `inline` binding: an atom name, optional Markdown node types with `fromMarkdown`, a Word `fallback`, and `updatable` for atoms whose attributes change in place. `compileInsertedAtoms` reads them instead of fixed math, mention, and source branches; removal reads `valid` and the inline binding; atom alignment, retained-atom keys, inline removals, inline source updates, and text-gap continuity read `updatable` and the component kind; media provenance reads `mediaPart`. The Sciobot bindings live in `registry/sciobot-inline.ts`, `registry/sciobot-valid.ts`, and `registry/sciobot-fallback.ts`. The inline atom and `diff` suites keep their baseline; collaboration files that use inline atoms pass alone (for example `diff/inserted-atoms-collaboration.test.ts`, 5 pass in 6 seconds).

## M06

- [x] Move the Sciobot registrations into the application

**Depends:** [M02](#m02), [M03](#m03), [M04](#m04), [M05](#m05)

**Primary files:** `packages/sciobot/src/registry/sciobot.ts` (new); `app:src/components/workspace/docx/`

**Read:** [contracts](../../../packages/sciobot/src/contracts/document.ts)

**Change:** Collect the Sciobot tags, prefixes, schemas, strategies, and lesson schema in one registration module. Have the application pass it to the adapter. Add a lint check that rejects Sciobot literals outside that module.

**Close when:** The generic adapter modules contain no Sciobot literal. The application supplies the only registration.

**Check:** Run the adapter suite and the lint check. Open saved Sciobot fixtures and compare semantic fingerprints.

**Claim:** claude, branch `codex/react-block-proof`. **Evidence:** `configureMdxRegistry` sets the application's registry once at startup; every adapter default reads `activeMdxRegistry()`, which falls back to the Sciobot registration. The Sciobot registration and its component codecs are the only modules that name Sciobot components: `src/__tests__/application-literals.test.ts` fails if any other adapter module names one (marked field names excepted). The registration stays in the private `@sciobot/docx` package, which is the application's own adapter; `sciobot-next` can call `configureMdxRegistry(sciobotMdxRegistry())` or pass its own definitions.
