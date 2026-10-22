# Application block decision receipts

R46 supplies independent decision storage. R15/R16 own the commands that resolve native revisions and application payloads together.

A property proposal can contain several block groups. Each group uses the stable pair `(proposalId, blockId)`. The proposal record stays unchanged when a group resolves. Never remove decided blocks by rewriting its shared JSON string. A real two-session probe concurrently retained different blocks from one proposal. Reconnect produced `shared-proposal:invalid-json`; `/tmp/codex-r46-shared-json-probe.log`.

## Stored representation

Use the existing `urn:sciobot:docx:review-groups:1` carrier with root `groups`. Initialize this part before room creation. Native-only records keep schema version 1. Each decision has a fresh `rg_` record ID and an empty native node label:

```json
{
  "schemaVersion": 2,
  "id": "rg_a_fresh_receipt_id",
  "proposalId": "shared-proposal",
  "blockId": "activity",
  "decision": "accept"
}
```

`decision` accepts `accept` and `reject`. Proposal and block IDs follow the existing block-ID rules. Record IDs follow the existing group-ID rules. The schema rejects extra fields, including canonical node IDs, paths, and Word revision addresses. Each receipt has a 4096-byte, four-level, 32-node bound. The existing carrier count, slot, and aggregate byte limits cover all records together.

Record data and labels never change. An exact same-ID write is idempotent. Replacing a receipt, reusing a removed native record ID, or tombstoning a receipt refuses. Native history owns undo and redo of receipt creation. Empty anonymous wrappers left by undo remain inert.

## Concurrent decisions

Each command creates a fresh receipt ID. Different block decisions never replace the same shared JSON node. Receipts for one group remain independent too. Equal dispositions read as accepted or rejected. Opposite dispositions read as `conflict`, with every receipt ID retained. No timestamp, actor order, or reconnect direction selects a winner. Removing one actor's receipt through undo reveals the remaining receipts. R22 retains explicit conflict reconciliation.

`readPayloadGroupDecisions` returns a pure map keyed by the serialized proposal/block pair. A missing key means pending. It refuses malformed carriers and unknown future records. It does not check whether an owner still exists. Retained receipts must survive owner deletion and history restoration. Decision commands must separately validate current ownership and native membership.

## Upgrade and scope

Existing property proposals and native-only group records need no rewrite. An existing shared room must already contain the R44/N21 carrier. Do not initialize a missing part separately in joined replicas. The first decision adds its receipt through the existing native package transaction.

Schema 3 and later remain opaque. Readers preserve their exact JSON text and refuse interpretation or replacement. Older private readers treat schema 2 as unsupported. Upgrade private review participants together before enabling group decision commands. Native participants can retain and transport these opaque nodes without a canonical-model or transport-format change.

R46 only supplies storage and pure disposition reads. Existing group readers ignore known receipt rows as native membership. Native revision reservation also skips these rows. R15/R16 must consume disposition state in group reads, property projections, and whole-proposal decisions. They must resolve exact native sites and payload changes in the same fresh transaction. Property fallback revision-ID retention also remains part of that integration. No production accept/reject command is enabled by this leaf.

The real-session fixture uses the EigenPal Pro License for transport testing only. Production modules import no evaluated implementation.
