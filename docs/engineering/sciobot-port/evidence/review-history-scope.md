# Shared native review history scope

R48 uses the established private owned-text protocol for native review history. R49 adds independent ownership of shared removal flags. Native revision acceptance removes formatting attributes and property snapshots. Each actor's removal intent must survive another actor's undo. A bare shared session does not provide that ownership.

The original same-group test passes when the owned protocol attaches before bootstrap. Its expected groups, label values, payload values, and native text stay unchanged. `/tmp/codex-r48-owned-probe.log` records both original cases passing. No native core or provider implementation changes are needed.

## Host lifecycle

Import the scope from the separate private collaboration entry. The main entry loads no Yjs implementation:

```ts
import { attachOwnedReviewHistory } from '@sciobot/docx/collaboration';

const history = attachOwnedReviewHistory(ydoc, { documentId });
// Create or join the native room only after attaching the scope.
const room = await createNativeRoom({ ydoc, documentId });
const session = history.bindSession(room.session);
// Supply session through the editor's native collaboration module.

// Release the editor and session before releasing their history scope.
editor.destroy();
room.destroy();
history.dispose();
ydoc.destroy();
```

`createNativeRoom` represents the host's injected native transport factory. This private entry imports no evaluated provider implementation. Every participant must install the existing owned protocol. The test transport uses the EigenPal Pro License.

The scope validates the room document ID, attachment timing, native protocol health, and disposal. Bind only the native session created over that scope's Y.Doc. Foreign document IDs and sessions already bound to another scope refuse. Repeated binding returns the same facade. Native functions retain their original receiver, including private class fields and frozen transport objects. The original session stays unchanged.

The facade exposes a local capability through a shared `Symbol.for` key. The main and collaboration entry bundles therefore agree on its identity. A module-local WeakMap would create separate identities in those bundles. The built-entry smoke check proves cross-entry use and disposal refusal; `/tmp/codex-r48-entry-smoke.log`.

`acceptReviewGroup`, `rejectReviewGroup`, `acceptAllReviewGroups`, and `rejectAllReviewGroups` check that capability before shared writes. A missing, disposed, destroyed, or unhealthy scope refuses before native operations, payload updates, or history entries. Solo review needs no collaboration scope. Existing native readiness, viewing, transaction, actor, selection, and history gates still apply.

## Native history behavior

The existing protocol owns text attribute intents and restored character lineage. It normalizes those attributes before native materialization publishes a canonical package. R49 applies that ownership rule to native removal flags. No extra undo step, hidden group filter, action replay, or alternate save pipeline is introduced.

Two actors can accept the same application group offline. Undoing either decision keeps the peer's native acceptance. Undoing the last remaining decision restores the original pending group and Word revision membership. Native Original/Proposed text and peer suffixes remain exact. The next ordinary undo removes the actor's prior typing, proving no extra normalization entry enters history.

The matrix covers both client orders, both reconnect receivers, both first actors, each actor's undo/redo, final remaining undo, edits inside the native fallback, owner deletion, remote geometry, room reopening, DOCX reopening, and saved convergence. The independent Word text oracle reads actual `w:ins` and `w:del` ancestry. Canonical revision sites must disappear while any accepted decision remains active. Restored pending sites must return when the final decision is undone.

## Recover review state after DOCX reload

Import `restoreReviewState` from `@sciobot/docx`. Pass the current canonical package and the current load's document ID, epoch, and revision. The pure read returns native groups, property/atomic plans rebound to current native owners, pending-review presence, and intrinsically reviewable group IDs. It reads no chat session, cached group list, or baseline document. It writes no content, metadata, display mode, shared state, or history.

Render review availability from `state.hasPendingReview`. Fully resolved decision records remain stored but produce no pending controls. Pending read-only imported groups remain visible; `state.reviewableGroupIds` excludes them. That list describes group ownership only. Native permissions, session readiness, private history, and complete capacity checks still apply when a decision runs. Use a fresh revision for each decision. Report a recovery failure instead of treating invalid or unsupported records as an empty review.

React hosts can subscribe once to native document changes and cache the result by immutable package identity. Selection-only changes need no recovery scan. The browser fixture uses `useSyncExternalStore` with that pattern. Its persistent fixture stores only DOCX bytes and an optional compatible native room update. A new load uses a new epoch. DOCX reopening starts a fresh native history; saved group identity does not transfer an old undo stack. Native room reopening also requires fresh client identities and the matching private protocol.

R21 has permanent real-browser reload cases for nested property/text review, actual keyboard input and native selection, mixed decisions, undo/redo, two sessions, remote geometry, owner deletion, and resolved-state reload without stale controls. Application review UI integration and external Word reconciliation remain separate leaves.

## Reconcile externally reviewed fallback text

`inspectExternalReview(pkg)` compares retained proposal records with their current native fallback owners. Exact captured old/new text endpoints identify clear external decisions after owned revisions disappear. Remaining owned revisions must still produce both captured views; partial or changed views report a conflict. Equal old/new labels provide no external decision evidence and remain pending. Already resolved immutable receipts keep their existing decision state.

`reconcileExternalReview(editor, revision, identity)` applies only clear property/atomic values and immutable decision receipts. Its report contains the validated observations. Its reconciled list names the metadata decisions applied. Read current review availability again through `restoreReviewState` after the native change event. A successful transaction can still report conflicts for other blocks. Preserve those blocks and show their diagnostic instead of treating the whole review as resolved.

The transaction writes application metadata only. Native Word content, unrelated revisions, external fallback edits, and opaque payload data remain intact. A single native undo restores the pre-reconciliation metadata and receipts. It does not undo the external Word review. Redo restores the same metadata decision. The complete staged removal/field admission and private history scope still apply. Group, bulk, and legacy property/atomic decisions share reconciliation's per-editor reentry guard.

Independent WordprocessingML fixtures exercise native acceptance/rejection results without calling the production decision implementation. They follow [Microsoft's Open XML revision examples](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document) and [Word text-revision rejection](https://learn.microsoft.com/en-us/office/vba/api/word.revision.reject). These fixtures do not establish desktop Word interoperability. V06 separately requires actual desktop Word saves, its version, and visual evidence.

## Bulk review decisions

`acceptAllReviewGroups` and `rejectAllReviewGroups` select all complete current groups from the live canonical package. They validate every member before one native transaction. One undo restores every pending group and its original identity. Empty calls add no history. Repeated stale calls and synchronous reentry refuse. Fresh document IDs, epochs, revisions, and surface checks prevent decisions against a replaced document.

Rejection retains manual content written between proposal waves. It resolves only the current pending native sites and application proposals. It never restores an earlier document snapshot. Both commands retain immutable proposal data and complete staged capacity preflight. Concurrent decisions, later unrelated groups, peer edits, owner deletion, history, geometry, and room/DOCX reopening have permanent source and installed-package checks.

## Plan removal of owned pending proposals

Import `planOwnedProposalRemoval` from `@sciobot/docx`. Pass the current canonical package, its revision token, and explicit host-retained ownership receipts. Each selected group needs its complete native revision addresses and property proposal/block pairs. Author names and nearby content do not establish ownership. Missing, changed, incomplete, or read-only groups refuse. The planner validates property and atomic fallback state through the existing decision preparation path.

The pure result contains scoped native rejection operations, immutable property rejection intents, and imported-address retention intents. It copies the selection in current document order. The detached result is frozen and bound to the supplied revision. It contains no canonical snapshots, editor handles, or callback. Planning changes no canonical content, metadata, references, display mode, or history. Native previews verify selected sites disappear while unrelated revision identities and authored attributes survive.

R24 supplies this pure plan. R25 retains live execution. A live executor must re-read and re-plan inside its current native gate. It must apply native inverse operations and immutable metadata decisions together with the replacement proposals. It must retain imported addresses before later native ID allocation. A returned plan alone does not change the document or provide an atomic replacement command.

## Native property history

Concurrent native property acceptance needs independent removal ownership. After both actors undo, ordinary map redo can retain a restored peer value. Correct active widths alone do not prove acceptance. Every `w:tblPrChange`, `w:tblGridChange`, and `w:tcPrChange` must disappear from the canonical document and saved Word file.

The private scope captures true-valued map flags only during the group's native transaction. Synchronous foreign transactions stay outside the local command's capture. It records each flag's retained Yjs map identity, field key, and missing/false baseline. Independent intent rows belong to that same native transaction. A repeated write to one flag uses one intent per command. The existing undo manager's original frame already names all inserted IDs. Before that manager first steps history, the scope adds the intent root through public `addToScope`. It matches the manager through its public tracked origins. It reads no provider field, schema name, or implementation property.

Undo removes only the stepping actor's intent rows. Redo restores those rows. Any active intent retains the flag. The final remaining undo restores the baseline. Untracked normalization transactions settle flags before native observers publish canonical state. Relevant flag changes and ownership changes trigger normalization; ordinary text edits skip flag scans. Incomplete update deliveries wait for their dependencies. Reversed normalization/intent delivery has focused proof.

This mechanism uses the pinned MIT Yjs peer. It requires retained shared map identities. Deleted targets remain inert. Restored map identities follow bounded local Yjs ancestry. Native table scenarios verify the actual retained record lifecycle. No whole review action is replayed, and no canonical revision is hidden from review reads or saves.

Native property tests cover both actor/client orders, both reconnect receivers, both redo orders, peer cell text, both Word width views, all six original property snapshots, ordinary typing/proposal history, row deletion, remote caret geometry, room snapshots, and DOCX reopening. Both original failing table cases retain their exact zero-pending-site assertion.

## Capacity admission

Before a shared decision commits, the adapter previews each native operation and each pure application package edit. It counts removed canonical node identities at every stage. This includes property snapshot descendants, opaque application XML, and transient nodes created and removed within the command. A retained native record has at most one removal flag. The count therefore bounds distinct flag intents for that command. Text-only or repeated removals can reserve unused headroom. Unchanged parts allocate no ID sets. The preview stays in the existing canonical pipeline and mutates no editor, shared document, or history.

The complete count travels through the private scope capability. `capture` requires that trusted count before it calls the native mutation callback. It reserves both potential target rows and retained intent identities. Undo hides an intent's value but retains its map key for redo. Admission counts those keys through the pinned Yjs map items. One participant therefore cannot spend another participant's redo headroom. Active parent reservations also reduce the headroom available to a nested command. Each distinct captured flag consumes one reserved slot. Return and exception paths release unused headroom. Invalid counts and insufficient capacity refuse before the callback runs. Incomplete Yjs updates also refuse; complete deliveries resume admission.

Admission uses the settled replica's state. It reserves the worst-case number of new target rows, including nodes whose records may already exist. A fresh decision therefore needs that headroom even when some target rows can be reused. Native undo and redo restore their existing intent rows through the original frame. They do not create another reservation or history step. Metadata-only decisions with no removed nodes can run at full ownership capacity. Host room admission remains separate from this local resource check.

The original two-flag capacity regression keeps its fixture and all unchanged-state assertions. Its call now supplies the complete two-flag count required by the private API. The first native flag, ownership rows, shared vector, and history stay unchanged on refusal. Exact target and intent boundary cases prove successful complete commands, ordinary undo/redo, unused reservation release, and nested admission. Real sessions prove refusal preserves Word XML, payloads, neighboring groups, canonical revisions, vectors, peer state, and prior redo. They also prove exact-budget acceptance, history, remote geometry, reconnect, room snapshots, and DOCX reopening.

## Compatibility and limits

The R51 runtime retains owned-text protocol 1 and advances private review-flag ledgers to protocol 2. The local capability uses `Symbol.for('sciobot.docx.review-history.4')`. Earlier capabilities refuse shared group decisions. String-field and string-sequence ownership share application protocol 1 in `sciobot-review-field-targets-v1` and `sciobot-review-field-intents-v1`. These are private application formats. The native provider's published collaboration version fields remain unchanged.

The flag protocol admits 50,000 targets and 100,000 retained intent identities. Inactive keys retained for native redo consume that intent limit. Field keys contain at most 256 characters. Target records contain at most 2,048 characters. Safe integer addresses, exact field sets, baseline values, protocol versions, and intent references have strict validation. Malformed room data poisons the scope and refuses later group writes. Attachment failures release partial ownership handles. Flag capture outside a native transaction refuses before writes.

R50 supplies complete local capacity preflight. The isolated application uses the verified R52 runtime described below.

Attach before creating native room content or loading a compatible owned-room snapshot. Attaching to an already-authored legacy room without the owned protocol refuses. All private participants must use the complete scope before group decisions. Hosts must gate that application protocol separately. Upgrade legacy live rooms through stopped-writer DOCX export and fresh-room bootstrap. This scope does not convert legacy history or support mixed private clients. Existing direct formatting consumers can retain `attachOwnedTextAttributes`; shared group decisions use the scoped session binding.

## R51 field and sequence ownership

Native table resolution reuses source geometry identities. Missing horizontal spans and vertical merge properties reuse their resolved marker identities. A missing owner container can reuse the removed cell's property container. Compacted columns retain a source column identity. Actor-bound resolution keeps original source paragraph identities through the existing R36 empty-retirement format. Transferred containers and markers detach from their original child slots before their cell disappears. Removed block controls also detach through ordinary canonical child-list primitives. Unrelated property identities retain their original child order. Canonical replay checks the resulting package.

The application scope owns string writes by opaque map address. It owns removal of existing string sequence entries by Yjs item identity. This preserves QName, attribute, width, and property-order decisions while another reviewer still owns the decision. It records no provider field names or schema. Sequence normalization removes restored entries while a removal owner remains active. Targets also retain whether an entry existed before its first owned write. Insertion and removal intents determine its presence. The final undo restores a preexisting entry at its retained neighbor anchors. Anchored order repair applies when that entry returns to present state. Later ordinary insertions retain their native order. An entry created within that undone gesture stays absent. This remains necessary after redo, because earlier normalization changes sit outside the native redo frame. Restored entries retain original item lineage. Repeated history does not replay document snapshots or reconstruct document content. The pinned Yjs search cache is invalidated before changes during observer delivery.

String writes outside review capture retain later ordinary edits. Existing string sequences also retain ordinary insertion and removal ownership. Newly created nested arrays retain native bootstrap ownership. Review commands reserve their complete budget before capture. Ordinary native writes retain the same per-record resource limits; host admission remains separate. Bootstrap remains outside field ownership. The scope excludes all ownership metadata roots. Restored sequence identities use the existing generic item lineage. No extra history entry is created.

The field ledger admits 50,000 combined scalar/sequence targets and 100,000 retained intent identities. Inactive keys retained for redo consume capacity. Each record admits at most 8 MiB of UTF-8 JSON. Root names and field keys admit 65,536 characters. Canonical previews reserve conservative headroom for names, attributes, namespace declarations, removed descendants, changed child lists, and package metadata. Each transient stage consumes headroom separately. Refusal occurs before the review mutation callback. Host-wide room admission and V05 performance checks remain separate.

### Upgrade private history scopes

Stop all writers before upgrading an existing private history-3 room. Export DOCX with its matching old runtime. Bootstrap a fresh room with matching history-4 participants and both initial synchronization directions. Do not load the old room snapshot into the upgraded scope. Preserve the old room separately if its history must remain available. DOCX export does not transfer room synchronization or undo history.

Gate the private application protocol before participants enter a room. All review participants need history 4 and flag protocol 2. A legacy protocol-1 header refuses attachment or live delivery. Detection of a losing legacy header persists an incompatible header before Yjs collects it. Reopening that mixed snapshot also refuses. This refusal supplements host admission; it does not provide mixed-client support or an in-place history migration. Native provider version checks remain necessary.

The isolated app uses guarded R52 runtime `92ec6f1e573fb7ef256b09f8b4889816c5c2f249`, adapter `sciobot-docx-25fc9fcca74e.tgz`, and rebuilt core `docx-editor.dev-core-84eceaf9ba22.tgz`. The three other rebuilt public archives retain their exact earlier bytes. Actual installed checks pass 178 cases across six files. They cover the owned-removal planner, leading-control paragraphs and independent pending properties, native property history, terminal-paragraph conflicts, and capacity admission. App units pass 53 cases; app/harness types pass 6646/2212 files. Production build, frozen install, archive bytes, and one-core/one-Yjs checks pass. Both browser files pass all 16 cases. Backup `refs/backup/codex-docx-app-r52-final` preserves the verified app artifact tree. Earlier R22 browser results remain recorded in the implementation notes. These checks do not close desktop Word V06, host admission, or V05 performance work.

## Dependent native operation stages

`TreeApplyOptions.stage` runs synchronously after initial native operations and before final package hooks. Its `TreeTransactionContext.currentPackage()` reads the intermediate working package. Staged native operations retain normal native validation, protection, actor allocation, cleanup, journals, and history. The context expires when its callback returns. Async callbacks and lifecycle operations refuse the complete transaction. Failed native operations or thrown stages publish nothing. A no-op stage retains the original package and redo state.

The isolated app uses R53 runtime `d3606be25437238394ab4f38387f00a29f499228`. Its core archive is `docx-editor.dev-core-d9dcb7e6e716.tgz`; its adapter archive is `sciobot-docx-8a1ae9b4b779.tgz`. Actual installed checks pass 183 cases across six files, including twelve stage-session cases. App checks pass 53 units, 16 browser cases, both type checks, and production build. Backup `refs/backup/codex-docx-app-r53-final` preserves the verified artifact tree. R25 still supplies ownership validation and atomic application replacement composition. Native stages alone do not establish that command.

### Owned proposal replacement

`replaceOwnedProposals(editor, selection, replacement, context)` replaces exact host-retained groups with a body paragraph text plan. The plan keeps the original revision token, but its coordinates refer to the document after selected inverses. Unsupported replacement kinds refuse. Ownership is explicit; matching author names do not grant ownership. The command returns frozen receipts for newly created groups.

Native inverse planning runs inside the review gate and before observed transaction capture. Its pure native previews must not enter the live journal. The live stage checks its exact starting package. A detached native coordinator validates intermediate text and reserves complete removal/field capacity, including cascades. Unselected revision fingerprints must survive replacement and final preview cascades. Selected inverses, immutable property decisions, retained imported addresses, replacement operations, and group metadata share one native publication and history unit.

R25's sixteen local and sixteen shared cases cover rollback, prior redo, save/reopen, concurrent peer edits/deletion, geometry, and capacity refusals. The actual installed command and shared inline lowering pass 102 cases across six complete files. R19 still owns later-wave preparation and orchestration. Public core and wire formats remain unchanged from R53.
