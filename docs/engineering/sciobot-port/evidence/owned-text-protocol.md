# Observed shared-text attributes

R39 supplies a private application boundary over MIT Yjs 13.6.32. R29 still owns tracked replacement placement.

Import `attachOwnedTextAttributes` from `@sciobot/docx/collaboration`. Supply the same pinned Yjs instance as the room. Attach before bootstrap, shared-text writes, or snapshot application. The document and React entries do not import this optional entry. Its bundle imports Yjs as an external peer. It imports no editor implementation.

```ts
const attributes = attachOwnedTextAttributes(ydoc, {
  onError: (error) => stopRoom(error),
});
// Start the application's room connection after checking its protocol contract.
// Before publishing a healthy room snapshot:
attributes.assertHealthy();
// After stopping editors and room observers:
attributes.dispose();
```

## Room contract

Every writer, reader, and export worker must use application protocol `OWNED_TEXT_PROTOCOL = 1`. The host must check this application contract before forwarding room updates. This helper does not implement transport negotiation. Native collaboration format equality alone does not establish application protocol equality.

Do not mix registered and legacy clients. Attach rejects previously authored rooms without this protocol. Unknown versions and malformed records fail explicitly. A failed handle rejects later writes. The host must stop the connection after `onError` reports a failure. Yjs cannot roll back an already received wire update. The boundary prevents native text observers from publishing malformed text.

To adopt the protocol for an old room, stop its writers, export the canonical DOCX, and create a fresh room. Attach the boundary before the new bootstrap. Do not replay legacy updates into that room. Retain the old room separately if its old synchronization identities or history are required. The application cutover gates retain the host integration and operational checks.

## Representation

Two named Y.Map roots contain the protocol data:

| Root | Data |
| --- | --- |
| `sciobot-owned-text-operations-v1` | `protocol: '1'` and immutable JSON operation records |
| `sciobot-owned-text-lineage-v1` | Restored character intervals mapped to their source identities |

Each operation records its insertion baseline or formatting changes, exact observed character ID intervals, and a causal ordering stamp. Unique `sciobot-owned:v1:<client>:<clock>` text attributes mark active operations. Undo and redo restore these markers through native history. Repair removes inherited markers from characters the operation never observed. Native encode passes cannot remove private markers as if they were document properties.

Before native observers run, the boundary computes attributes from each character's active insertion baseline and owned formatting operations. Concurrent independent fields combine. Conflicting fields use the causal stamp, then client ID and clock. A later observed write wins. Repair uses the original Yjs methods and a separate origin. Repair does not create a user undo unit.

Yjs undo creates new character identities. The boundary reads local redo links and publishes interval lineage. Receivers use that lineage before native observers read restored text. Named roots initialize before updates arrive, including on fresh joining replicas. Native serialization ignores the private metadata. Tests check every exported XML part.

## Lifecycle and limits

Installation wraps Y.Text authoring methods once. A WeakMap enables the wrappers only for attached documents. Unregistered and disposed documents retain ordinary Yjs behavior. Disposal removes the before-observer hook. The host must dispose the handle after stopping the room.

The boundary preserves UTF-16 positions, tab characters, embedded values, and native delta trailing-newline behavior. Adjacent equal repairs form one formatting range. This preserves supplementary characters during repair. It detaches attribute data and validates complete deltas before authoring their operations. Native embed payload semantics remain Yjs-owned.

Attribute input allows 8 MiB, 1,000,000 JSON nodes, and depth 32. An operation record allows 16 MiB and 100,000 observed intervals. Invalid accessors, cyclic data, unsafe property keys, malformed ranges, and unsupported schema fields fail explicitly. Lineage records allow 1 KiB. Historical scans and per-character normalization still require the V05 performance gate. These correctness tests do not establish large-document performance.

## Verification

```bash
bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/collaboration
bun packages/sciobot/test/diagnostics/yjs-format-boundaries.ts
bun packages/sciobot/test/diagnostics/yjs-owned-format-proof.ts
```

The legacy diagnostic must retain its original failure. The independent owned-format diagnostic demonstrates the ordering prerequisite. Production tests use the exported boundary.

The MIT-only cases cover both fixed client orders and reconnect receivers, observed ranges, independent field conflicts, insertion history, repeated restored parents, snapshots, tabs, embeds, supplementary characters, strict parsing, disposal, and observer failure handling.

Twenty native cases use real sessions through the unchanged evaluation fixture. Sixteen cover plain or italic peer suffixes, both decisions, both client orders, and both reconnect receivers. They verify exact Original/Proposed text, proposal and decision history, native observer publications, remote selection geometry, raw-room continuation, saved convergence, DOCX reopen, and metadata-free XML. Four cover row deletion during peer formatting, restored peer proposals after deletion undo, redo, and saved convergence.

The evaluation fixture uses the EigenPal Pro License. Production protocol source imports neither that fixture nor its transport implementation. The fixture enables this boundary explicitly. R29's original replacement tests remain unchanged and unchecked until native replacement placement integrates both requirements.
