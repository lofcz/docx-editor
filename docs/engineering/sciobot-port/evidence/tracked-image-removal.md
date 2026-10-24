# Tracked image removal and late resource references

Branch: `codex/d30-tracked-images`, after `ffaeea1e`.

## Removal checkpoint

Managed image removal uses the existing native whole-control deletion. Presence ownership retains the payload and resource provenance until review. Acceptance retires the exact payload and removes only orphan image resources. Rejection restores the complete original picture and its link.

An image-only hyperlink needs no second unlink proposal over its deleted control. The stage consumes unlink only when every link child is an exact deletion site produced by its inline removal operations. Native acceptance removes the emptied link. Native rejection restores its original metadata. Links containing foreign sites or other children retain native refusal rules.

The affected source batch passes all 67 cases and 2,735 assertions across five files. This includes linked and unlinked pictures, a surviving image sharing the removed image's media, independent media, exact native acceptance and rejection, exact surviving picture fingerprints, payload retirement, creation history, decision history, both actor orders, offline peer text, geometry, Word reopening, and cold room reopening. Image birth oracles also require Word control identities.

The built SDK and native core pass all 16 removal cases and 1,184 assertions. Neither source nor compiled tests change their native or media oracles.

Logs: `/home/lofcz/.cache/sciobot-docx-native/d30-image-removal-*`.

## Critical remaining failure

`image-removal-late-reference.test.ts` exposes data loss under both client orders. Alice proposes image removal and accepts while disconnected. Bob independently creates a new drawing referencing the original image's existing resource, without adding bytes. After reconnect, the canonical drawing survives and matches the independent native expected tree, but the image media part is absent.

The fixture verifies the peer copy before reconnect and derives expected output through independent direct removal from that peer's package. Expected media remains referenced and must survive. Both cases fail the complete media comparison. The assertions remain unchanged. Native drawing relationships, decision history, and cold room reopening follow that assertion and remain unverified for this case.

The transport retires binary descriptors by clearing their digest. It also clears relationship identities. The retained node structure lets the late drawing survive, but does not retain enough readable resource information to restore its bytes. Fix resource lifetime using actual live references; do not keep every removed asset visible or weaken exact media checks. A format change, if required, needs its own compatibility decision, migration guide, and release note.

D30 and R19 remain open. Do not claim a complete image removal gate or production cutover from the passing checkpoint cases.

## Implementation investigation

The next repair must recover retired resources only when live canonical drawing references require them. An immutable original digest on each binary entry can preserve byte identity while leaving the existing retirement write unchanged. Relationship holders retain the original identity as their inner map key. Recover the exact native target and MIME declaration, then project matching relationships parts and bytes together.

Cache reference evidence across unchanged immutable nodes. Ordinary text edits must keep the existing zero relationship-decode behavior. Verify removal with no remaining references, surviving shared references, late offline copies, removal of the copied drawing's owner, both histories, and cold reopening. Keep the original room admission and memory limits.

Blindly republishing resources on every observed drawing edit can leave orphan bytes when a concurrent edit removes that drawing's owner. Such a write-only workaround does not prove the required lifetime semantics.
