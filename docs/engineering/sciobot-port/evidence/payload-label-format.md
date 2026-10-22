# Immutable application label updates

R47 fixes the stored application label. It does not change native Word fallback text or data bindings.

Each application record keeps its created `label` and `data` text. Later label updates append independent `labelValue` children:

```xml
<node id="scio_middle">
  <label>Stored middle</label>
  <data>...created payload JSON...</data>
  <labelValue version="1" value="&quot;New middle&quot;"/>
</node>
```

The surrounding namespace remains `urn:sciobot:docx:blocks:1`. The `value` attribute contains one JSON string. JSON escaping preserves tabs, carriage returns, newlines, quotes, and multilingual text through XML attribute serialization. Empty strings remain valid. Each label uses the existing two-million-character limit of 2 _ 1024 _ 1024 UTF-16 units.

Each update owns a fresh native child identity. That identity does not become an authored XML ID, JSON field, canonical path, or stored group address. Existing rows never change. The last authored row in native document order selects one complete value. Concurrent values never merge character by character. Undo removes only its actor's row. Redo restores that row without replacing the peer's row. Empty native wrappers left by undo remain inert.

A single shared scalar attribute also failed the required history test. After Alice's undo/redo and Bob's undo, the effective label returned to its created value. Independent immutable children preserve Alice's earlier value. The failed prototype is replaced; `/tmp/codex-r47-focused.log` records its failure. The final label history suite covers both client orders, both reconnect receivers, equal/different values, and both legacy and prepared records.

The reader bounds each record to 10,000 authored label updates, 50,000 child slots, and eight MiB of encoded label values. It validates every authored row. Known rows require version `1`, exactly one JSON string value, no extra attributes, and no children. Invalid XML characters, malformed JSON, duplicate fields, and oversized values refuse. Unknown label versions refuse interpretation and writes while preserving source data. Unsupported future payload versions retain their own label semantics and opaque metadata.

Legacy records need no conversion. Their created label remains effective until an update exists. A label update uses the existing record and carrier. No new shared part, sentinel, eager initialization, or alternate document pipeline is introduced. Unchanged writes preserve package identity. Payload attribute registers and opaque record metadata remain intact.

Application blocks use tags to locate payloads. Their native generated Word text remains separate and changes through native operations. Created `label` text and existing Word data bindings stay unchanged. Bound property controls retain the existing explicit native reconciliation requirement. Foreign stores, foreign binding XML, and native Word parts remain intact. R47 does not claim new bound-control editing support.

Older private application readers ignore `labelValue` children. Upgrade private participants together before enabling these writes. Native participants can preserve and transport the opaque nodes without a native format change. No published package, public document API, provider implementation, or shared transport schema changes.

R47 closes independent label storage and label history. The stronger same-group acceptance test also checks canonical native revisions. Local undo can restore a native insertion wrapper while the peer's accepted receipt remains active. R48 retains that separate history requirement. Do not hide the restored revision or treat label correctness as proof of complete group acceptance.
