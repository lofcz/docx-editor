# Tracked image removal and late resource references

Branch: `codex/d30-tracked-images`, after `ffaeea1e`.

## Removal checkpoint

Managed image removal uses the existing native whole-control deletion. Presence ownership retains the payload and resource provenance until review. Acceptance retires the exact payload and removes only orphan image resources. Rejection restores the complete original picture and its link.

An image-only hyperlink needs no second unlink proposal over its deleted control. The stage consumes unlink only when every link child is an exact deletion site produced by its inline removal operations. Native acceptance removes the emptied link. Native rejection restores its original metadata. Links containing foreign sites or other children retain native refusal rules.

The affected source batch passes all 67 cases and 2,735 assertions across five files. This includes linked and unlinked pictures, a surviving image sharing the removed image's media, independent media, exact native acceptance and rejection, exact surviving picture fingerprints, payload retirement, creation history, decision history, both actor orders, offline peer text, geometry, Word reopening, and cold room reopening. Image birth oracles also require Word control identities.

The built SDK and native core pass all 16 removal cases and 1,184 assertions. Neither source nor compiled tests change their native or media oracles.

Logs: `/home/lofcz/.cache/sciobot-docx-native/d30-image-removal-*`.

## Reproduced resource-loss baseline

`image-removal-late-reference.test.ts` exposes data loss under both client orders. Alice proposes image removal and accepts while disconnected. Bob independently creates a new drawing referencing the original image's existing resource, without adding bytes. After reconnect, the canonical drawing survives and matches the independent native expected tree, but the image media part is absent.

The fixture verifies the peer copy before reconnect and derives expected output through independent direct removal from that peer's package. Expected media remains referenced and must survive. At `e861d989`, both cases fail the complete media comparison. The assertions remain unchanged. Native drawing relationships, decision history, and cold room reopening follow that assertion and remain unverified for this case.

The transport retires binary descriptors by clearing their digest. It also clears relationship identities. The retained node structure lets the late drawing survive, but does not retain enough readable resource information to restore its bytes. Fix resource lifetime using actual live references; do not keep every removed asset visible or weaken exact media checks. A format change, if required, needs its own compatibility decision, migration guide, and release note.

D30 and R19 remain open. Do not claim a complete image removal gate or production cutover from the passing checkpoint cases.

## Implementation investigation

The next repair must recover retired resources only when live canonical drawing references require them. An immutable original digest on each binary entry can preserve byte identity while leaving the existing retirement write unchanged. Relationship holders retain the original identity as their inner map key. Recover the exact native target and MIME declaration, then project matching relationships parts and bytes together.

Cache reference evidence across unchanged immutable nodes. Ordinary text edits must keep the existing zero relationship-decode behavior. Verify removal with no remaining references, surviving shared references, late offline copies, removal of the copied drawing's owner, both histories, and cold reopening. Keep the original room admission and memory limits.

Blindly republishing resources on every observed drawing edit can leave orphan bytes when a concurrent edit removes that drawing's owner. Such a write-only workaround does not prove the required lifetime semantics.

A resource recovery pass must also restore the exact `.rels` XML projection and MIME override. Preserve original relationship-row extensions and relationships-root metadata, including when removing the final relationship removes its directory entry. Keep immutable initial resource descriptors at seeding/put time, rather than adding extra scalar writes to each retirement history step. Any additional shared descriptor fields require a schema upgrade and coordinated room migration.


## Reference-dependent recovery checkpoint

The evaluated transport retains immutable binary digests and original MIME declaration modes at resource creation. Relationship parts retain their original root identity after retirement. Materialization restores retired assets only when a live canonical relationship reference needs them. Removing the drawing owner releases the recovered resources. No new binary or relationship publication pins a deleted owner.

Relationship XML recovery preserves original root and row extension attributes. The reader follows changes to retained XML records and invalidates only the affected recovery cache. Text-only edits reuse recovered bytes without new blob reads or relationship decodes. Full and incremental materialization, Word reopening, and cold room reopening retain the complete tested native package.

The source batch passes 88 tests and 2,616 assertions across ten files. This includes image births, removals, both offline client orders, resource fidelity, invalid retained descriptors, missing blobs, format refusal, relationship journals, and image publication cost. The compiled SDK and core pass all 43 image cases and 2,367 assertions. An additional compiled run verifies both actors' histories, including resource release after undoing Bob's late drawing and restoration after redo: two cases and 142 assertions. Pro type checks pass 2,327 source files and 1,876 Vue-profile files.

Logs: `/home/lofcz/.cache/sciobot-docx-native/d30-retained-resource-{batch,compiled,peer-history,types,policy}.log`.

The evaluated format changes to `docx-collaboration:1.7.7.1`. Previous rooms refuse admission without writes. Upgrade instructions require verified old-client export and fresh room seeding. No saved room is migrated here. Pro changes remain evaluation-only under the EigenPal Pro License; production adapters import no Pro transport.

D30 and R19 remain open. The broader composer baseline still has twelve unresolved failures. This checkpoint does not establish full composer or application parity.
