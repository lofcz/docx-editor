# Property review record format

The private adapter stores pending property records beside native application payloads. Each record occupies one custom XML node.

The existing store uses namespace `urn:sciobot:docx:blocks:1` and root `blocks`. A property owner already has a payload in this store. The record writer requires that existing store.

Block payload nodes use the `scio_` prefix. Review nodes use the separate `scio-review_` prefix. Block readers exclude review nodes. Review readers validate them through their own bounded schema.

For proposal `proposal-one`, the node ID is `scio-review_proposal-one`. Its label is empty while pending. Its data contains this JSON shape:

```json
{
  "schemaVersion": 1,
  "id": "proposal-one",
  "author": "ScioBot",
  "date": "2026-10-19T09:00:00.000Z",
  "predecessors": [],
  "blocks": [
    {
      "blockId": "activity",
      "type": "lesson_activity",
      "changes": [
        {
          "name": "duration",
          "before": { "present": true, "value": 10 },
          "after": { "present": true, "value": 15 }
        }
      ]
    }
  ],
  "nativeRevisions": [
    {
      "partName": "/word/document.xml",
      "id": "5",
      "author": "ScioBot",
      "date": "2026-10-19T09:00:00.000Z"
    }
  ]
}
```

`present: false` has no value field. Null, zero, false, and absence remain distinct. Each block has unique changed property names. Duplicate block IDs, no-op changes, extra schema fields, and malformed native addresses refuse. An optional `extension` object retains opaque JSON data.

Records contain stable application block IDs. They contain no saved canonical control IDs or mount revision tokens. The reader resolves each block against the current native block index. A missing, duplicated, or type-changed owner causes an explicit failure. Reading never recreates a block.

`predecessors` names proposals observed before this proposal. The reader orders present records through these edges. Missing predecessors can represent earlier resolved records. Cycles refuse. Concurrent proposals for one property lack causal ownership and refuse. Property value continuity also remains mandatory. XML sibling order and wall clocks cannot establish proposal order.

`nativeRevisions` retains complete part-scoped Word revision addresses. The schema preserves optional address dates exactly. Storage tests prove address retention, not fallback revision creation or resolution. R12 validates and executes those operations through its guarded property transaction.

Removal changes the node label to `removed`. The record data and native field identities stay intact. Removed IDs remain reserved. This avoids deleting a peer node when materialized sibling order differs from shared order. Undo restores the pending marker. An unknown-version record cannot be overwritten or removed through this API.

Native peer-preserving undo can leave an empty anonymous node. It has no attributes, children, ID, or record. Readers preserve it as inert XML. A nonempty node without an ID remains malformed.

The limits are 2 MiB per record, depth 64, and 50,000 JSON nodes. One record contains at most 1,024 blocks. Each block contains at most 128 property changes. A record contains at most 4,096 native revision addresses and 1,024 predecessors. The review store permits 1,024 record nodes and 8 MiB of record data. The enclosing application store has its own limits.

Future schema versions retain their exact raw record data. Known-version readers refuse to derive a view from an active future record. Native stories and opaque data remain preserved.

Native review reads automatically include stored pending records. The caller can add detached pending plans through `pendingProperties`. Duplicate IDs refuse. Each committed or working read uses one native package revision and its corresponding review layout.

The storage helpers return immutable package candidates. Live writers must use the native gate and one transaction. R12 owns property transactions, attributed fallback revisions, and decisions. R14-R16 own persistent review groups and atomic group decisions.

R12 records use extension marker `executor: "sciobot-properties-v1"`. Their `fallbackLabels` array contains one entry per block. Each entry retains `blockId`, the stored labels `before` and `after`, and generated paragraph text `textBefore` and `textAfter`. These strings preserve execution evidence without saved canonical node IDs. Changed generated text requires retained native revision receipts. Canonical attributes stay committed until acceptance. Rejection leaves them intact. Group and external Word reconciliation remain separate requirements.

R13 atomic records use the same bounded schema with `executor: "sciobot-atomic-v1"`. Each block contains exactly one change named `$payload`. Both values are present and contain complete supported block payloads. IDs and payload types remain equal. The reader interprets this slot as an empty payload path. It never writes an attribute named `$payload`. The executor validates the atomic diff strategy and opaque MDX before creating a proposal. Stored causal edges also protect atomic replacements from conflicting concurrent ownership.

The `fallbackLabels` extension retains the same label and text evidence as R12. Canonical payloads retain original assets until acceptance. Review reads select the complete old or proposed payload. Acceptance writes the complete proposed payload and selected label. Rejection retains original payloads and labels. Both decisions resolve only retained native revision addresses. Missing, partial, externally resolved, or foreign receipts refuse. Removal reserves the record ID and retains both payloads for peer-preserving undo. R14/R19 retain overlapping groups and later-wave reconciliation. R22 retains external Word decisions.
