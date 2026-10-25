# List toggle and removal checkpoint

N08 remains open. This checkpoint implements source toggle behavior within the private list adapter.

A collapsed style conversion changes the original same-level, same-style sibling run. It skips deeper blocks and stops at a lower level, an unindented block, or a different style at the same level. An expanded conversion changes only selected paragraphs. These conversion boundaries follow `getListSiblings`; heading sequence boundaries remain separate.

Toggling an active numbered style removes list membership and retains one fewer source indentation step. The native writer uses explicit zero numbering to disable inherited membership. It preserves trailing indentation, foreign attributes, namespace bindings, and opaque children. Tracked membership and indentation changes use separate native staging steps. This lets the detached preview reserve the removal flags for intermediate property snapshots.

An exact two-editor rejection test exposed indentation holder replacement as a second defect. The native property writer now preserves the existing Word `ind` identity during attribute changes. Both client orders restore original leading geometry while retaining the peer's trailing geometry. Each author keeps independent undo and redo. This uses the internal field-2/flag-3 prototype; default production admission remains unchanged.

Removing a numbered item splits the following counter into a new instance. The adapter copies the complete imported `num` descriptor, including abstract references, overrides, foreign attributes, and nested formatting. It replaces a matching start override instead of adding a duplicate override. Existing abstract definitions and counter descriptors remain unchanged. Native counter allocation and template copying use separate staging boundaries. This lets the reservation preview include removals of newly allocated temporary children.

The focused batch passes 30 cases and 332 assertions across seven files (`/tmp/n08-toggle-reset-integration.log`). It includes both shared client orders, disconnected peer text, direct writes, both decisions, independent histories, native saves, and exact Word-tree comparisons. The affected regression batch passes 255 cases and 1761 assertions across twelve files (`/tmp/n08-toggle-affected-regression.log`). It covers native property/revision operations, numbering identities, private plan validation, sequence joining, checklist review, and shared indentation.

The private build verifies 328 modules with core external. Browser type checking passes 3672 files across two projects (`/tmp/n08-toggle-browser-types.log`). This checkpoint does not establish polite restart parity, default production ownership, or portable pending-DOCX decisions.

Remaining source behavior includes complete outdent counter resets and polite restarts. The removal checkpoint does not establish complete multilevel or imported restart parity. The final N01 capture and browser/shared completion gates still apply.

The strengthened peer-indentation cases pass both client orders after pending-room rehydration. They verify exact saved/reopened Word trees, independent histories, matching caret geometry, remote selection addresses, and painted remote carets (`/tmp/n08-toggle-cold-peer-proof.log`: two cases, 60 assertions).

A nested-prefix regression reproduced `2.1.` becoming `1.1.` when the child tail moved to a new counter. The adapter now reads parent counter values through Apache core’s existing counter engine. It carries bounded ancestor starts into the cloned instance. Exact native descriptor and main-part oracles cover direct edits, both decisions, creation/decision history, disconnected peer text, peer history, both client orders, and saved output. The combined batch passes 47 cases and 582 assertions across nine files (`/tmp/n08-ancestor-final-batch.log`). Private types pass 2417 files (`/tmp/n08-ancestor-types.log`).

The initial Chromium run times out during `page.goto`, before list actions. A readiness probe also times out on Vite module responses. Verification now uses the fixture’s existing stable-server mode. The stable server completes the list-entry readiness probe in 14.9 seconds. The focused repeat writes its results to `/tmp/n08-toggle-stable-browser-proof.log`. The original 60-second test deadline remains unchanged. The source-oriented fixture timed out during page loading after 386 module requests. A built-core fixture run then passed all three new Chromium cases in 18.1 seconds (`/tmp/n08-built-browser-proof.log`).

The final private build verifies 329 modules with core external (`/tmp/n08-ancestor-private-build.log`). Native core build, including compiled editor styles, also passes (`/tmp/n08-indent-core-build.log`).
