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

- [ ] Define the embedder MDX component contract

**Depends:** [B01](02-react-blocks.md#b01), [B22](02-react-blocks.md#b22)

**Primary files:** `packages/sciobot/src/registry/` (new); `packages/sciobot/src/index.ts`

**Read:** [payload codec](../../../packages/sciobot/src/blocks/payload-opc-codec.ts), [payload schema](../../../packages/sciobot/src/blocks/payload-schema.ts), [block tags](../../../packages/sciobot/src/projection/block-tags.ts), [MDX tags](../../../packages/sciobot/src/projection/mdx-tags.ts)

**Change:** Add the `MdxComponentDefinition` type and a frozen registry value with the fields listed at the top of this page. Validate names, prefixes, and namespaces at registration. Refuse duplicates and prototype keys.

**Close when:** A registry can be built from definitions. Invalid, duplicate, or conflicting definitions refuse with a stable reason.

**Check:** Add focused registry construction tests, including hostile names and prefixes.

**Claim:** Unclaimed. **Evidence:** Pending.

## M02

- [ ] Drive the MDX projection from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/projection/mdx.ts`; `packages/sciobot/src/projection/mdx-tags.ts`; `packages/sciobot/src/projection/block-tags.ts`

**Read:** [MDX projection](../../../packages/sciobot/src/projection/mdx.ts), [MDX tags](../../../packages/sciobot/src/projection/mdx-tags.ts)

**Change:** Replace `DOMAIN_BLOCK_TAGS` and the `CUSTOM_*_TAGS` constants with registry reads passed through `projectBlocks` and `projectMdx`. Keep the opaque `shared_mdx_*` lane for allowed but unregistered tags.

**Close when:** The projection reads no module constant for application tags. Existing projection fixtures produce the same output with the Sciobot registration.

**Check:** Run the projection and MDX suites. Add one fixture with a non-Sciobot component.

**Claim:** Unclaimed. **Evidence:** Pending.

## M03

- [ ] Drive payload storage from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/blocks/payload-opc-codec.ts`; `packages/sciobot/src/blocks/payload-part.ts`

**Read:** [payload codec](../../../packages/sciobot/src/blocks/payload-opc-codec.ts), [payload part](../../../packages/sciobot/src/blocks/payload-part.ts)

**Change:** Take the content-control tag prefix and the custom XML namespace from the registry instead of `BLOCK_TAG_PREFIX` and `PAYLOAD_NAMESPACE`.

**Close when:** Documents saved before this change open with identical payloads. A second registration stores its payloads under its own namespace.

**Check:** Run the payload suites against saved fixtures. Add a save and reopen case for a second namespace.

**Claim:** Unclaimed. **Evidence:** Pending.

## M04

- [ ] Drive diff identities and strategies from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/diff/domain-strategies.ts`

**Read:** [domain strategies](../../../packages/sciobot/src/diff/domain-strategies.ts)

**Change:** Replace the `STRATEGIES` table and `domainDiffStrategy` lookups with the registered strategy. Keep the generic strategy for unregistered components.

**Close when:** The diff modules hold no component-name literal. The existing diff suites pass with the Sciobot registration.

**Check:** Run the diff suites. Add a planner case for a non-Sciobot component.

**Claim:** Unclaimed. **Evidence:** Pending.

## M05

- [ ] Drive inline atoms from the registry

**Depends:** [M01](#m01)

**Primary files:** `packages/sciobot/src/projection/math.ts`; `packages/sciobot/src/projection/mention.ts`; `packages/sciobot/src/react-blocks/inline-host.tsx`

**Read:** [inline host](../../../packages/sciobot/src/react-blocks/inline-host.tsx), [mention projection](../../../packages/sciobot/src/projection/mention.ts)

**Change:** Replace the fixed `math`, `mention`, and `source` atom kinds with registered inline components. Keep equations native.

**Close when:** A registered inline component projects, measures, renders, and deletes as one native atom.

**Check:** Run the inline atom suites. Add a two-session case for a non-Sciobot inline component.

**Claim:** Unclaimed. **Evidence:** Pending.

## M06

- [ ] Move the Sciobot registrations into the application

**Depends:** [M02](#m02), [M03](#m03), [M04](#m04), [M05](#m05)

**Primary files:** `packages/sciobot/src/registry/sciobot.ts` (new); `app:src/components/workspace/docx/`

**Read:** [contracts](../../../packages/sciobot/src/contracts/document.ts)

**Change:** Collect the Sciobot tags, prefixes, schemas, strategies, and lesson schema in one registration module. Have the application pass it to the adapter. Add a lint check that rejects Sciobot literals outside that module.

**Close when:** The generic adapter modules contain no Sciobot literal. The application supplies the only registration.

**Check:** Run the adapter suite and the lint check. Open saved Sciobot fixtures and compare semantic fingerprints.

**Claim:** Unclaimed. **Evidence:** Pending.
