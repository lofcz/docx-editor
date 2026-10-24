# Committed text beside pending insertions

D30 and R19 remain open. This checkpoint fixes the native rejection failure for peer text inside born atom paragraphs.

Untracked insertion previously changed the existing text leaf inside `w:ins`. The peer text then belonged to that insertion. Native rejection correctly removed the wrapper and therefore removed the peer text too.

The canonical writer now places committed text independently of pending insertion ownership. It reuses the existing native insertion splitter. The committed path creates no temporary revision and consumes no revision ID. Attributed typing keeps its author and timestamp rules. Text, tabs, manual breaks, page breaks, and page-field payloads retain their model widths. Builders run once; emitted native nodes carry the actual fresh-text capture.

When a stable hyperlink, unlocked control, custom XML container, or smart tag sits inside the pending insertion, the committed path moves insertion ownership inside that container. The container keeps its original identity, properties, imported attributes, and namespace metadata. Existing attributed insertion refusal rules remain unchanged. Imported generic SDT property holders remain structural metadata. The native inline depth limit remains unchanged.

Verification uses frozen source and unchanged expected Word, payload, identity, history, and reopen assertions:

- The combined source batch passes 352 cases across 15 files and 3498 assertions. It includes C01, direct execution, composers, tracked author behavior, inline wrappers, original owned-presence histories, retained atom births, and new paragraph atoms. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-committed-integration.log`.
- The strengthened shared review file passes eight cases and 244 assertions. It covers both client orders, concurrent offline typing and rejection, exact pending and decided native fingerprints through decision undo/redo, payload retirement, healthy ownership, and room reopening. Log: `d30-committed-review-history-final.log` in the same directory.
- The compiled core/private SDK/evaluation transport passes 43 cases across three unchanged fixture bodies and 617 assertions. It includes 16 native metadata goldens, 19 retained atom cases, and all eight new-paragraph shared cases. Log: `d30-committed-sdk-final.log`. The harness changes import targets only. All engine imports resolve to one built core. The evaluation transport compilation covers its collaboration entry, not a full Pro package release build.
- Core, i18n, fonts, Markdown exporter, and private SDK package builds pass. Core types cover 1610 files; Pro source types cover 2318 files; private types cover 2254 files.
- The Markdown declaration build initially resolved the main checkout's stale core through a borrowed dependency link. A dependency overlay owned by this worktree now resolves core, fonts, and i18n to these local artifacts. The verified exporter build passes without installation or shared dependency changes.
- Core API checking passes 15 entries with zero errors and no snapshot update. It reports 2489 warnings and 391 forgotten exports. No baseline is recaptured.
- Scoped lint, max-lines caps, compatibility policy, formatting, and whitespace checks are recorded at the checkpoint. The compatible decision introduces no shared format field or reader-rule change.

The implementation remains in the Apache core path. Test-only transport uses the EigenPal Pro License. Production private adapter modules import no Pro implementation. Tracking settings do not change automatically.

Remaining D30 work includes tracked image resource ownership, new-cell atom ownership, arbitrary retained atom hyperlink placement, native OMML deletion, and final complete integration gates. The application remains read-only. No TODO closes from this native repair alone.
