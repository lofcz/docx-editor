# Retained inline atom mark changes

Runtime `956ba1e9` adds explicit native mark plans for unchanged inline sequences containing registered atoms. Source mapping resolves each atom to its exact native range. Bold, italic, strike, and code changes retain the original controls, payloads, drawings, and embedded media. Existing native property planners preserve other authored properties. Source attribute order remains incidental under the existing atom key.

The direct executor now validates direct formatting as an untracked operation. It no longer validates a temporary tracked operation before removing its attribution. This matters because native content-lock policy deliberately refuses tracked formatting inside locked content. Direct previews accept the host's explicit formatting policy. The live editor must independently enable that same policy. Neither option changes stored locks. Two refusal cases verify zero publications when either side omits the opt-in.

Core inline drawing records now carry optional owning run formatting from the existing resolved style cascade. The Markdown exporter represents drawing emphasis and code through structural wrappers outside the replaceable image source range. Images keep their original native source mapping. This avoids flattening code-styled image syntax into literal backticks. The change has a minor core changeset and a compatible collaboration decision. The new fields belong only to derived layout records.

## Focused evidence

At `956ba1e9`, the two new suites pass 22 cases and 666 assertions. Sixteen local cases cover four atom kinds and four mark kinds. Each case adds and removes its mark. It compares the complete canonical tree with an independent native formatting oracle, tests undo/redo, saves and reopens DOCX, and replans to no operations. Media resolution is forbidden during these retained edits. Native IDs, payload records, and existing media stay exact. Code-style creation can add the required style definition, relationship, and content-type entry.

Four real-session cases cover both peer orders and both formatting writers. They format retained math, mentions, sources, and an image. A disconnected peer edits the same paragraph. Exact native trees survive reconnect and both actors' histories. Remote selection geometry agrees. DOCX and room reopening retain native output, payloads, and a no-op target plan. The EigenPal Pro License transport remains evaluation-only.

Commands:

```sh
bun test --preload ./scripts/test/preload-dom.ts packages/sciobot/src/diff/retained-atom-marks.test.ts packages/sciobot/src/diff/retained-atom-marks-collaboration.test.ts
```

Logs: `/tmp/d30-retained-atom-marks-focus.log`, `/tmp/d30-retained-atom-marks-existing-focus.log`, and `/tmp/d30-retained-atom-marks-native-exporter.log`.

The affected private check passes 183 existing cases across 15 files with 5,857 assertions. It includes C01, retained/link/code/image materialization, direct execution, property suggestions, and recomposition. The exporter/native drawing check passes 284 cases across 37 files with 1,601 assertions. Its 500-page export remains within the original resource limits.

D30 remains open. Changed hyperlink ownership, changed atom sequences, image removal/replacement, new nested cell fragments, tracked atom ownership, and installed application parity still need their stated evidence. No leaf closes here.

## Built consumers and required checks

Fresh core, exporter, and private SDK builds pass. The private package boundary verifies 257 modules with core external. The same 22 original fixtures pass against compiled packages with 666 assertions. The runner changes imports only:

```sh
bun packages/sciobot/scripts/check-code-inline-controls-compiled.mjs retained-atom-marks.test.ts retained-atom-marks-collaboration.test.ts
```

Core types pass across 1,602 files. Exporter types pass across 1,531 files. Private SDK types pass across 2,191 files. Core and exporter API extraction includes the new drawing formatting descriptor. Their explicit type exports satisfy the forgotten-export policy without changing its baseline. Root formatting and lint pass. Compatibility policy passes 185 decisions and two inherited format fields. The earlier table-holder checkpoint also passes all 56 compiled code-control cases and 1,267 assertions.

The backlog check still fails its unchanged production-inventory gate. The application revision is `073cb22447c6ac6a3699958e27bb1f3b8661bd22`; the pinned revision is `5650bdf550b7987174d1b61fe2a925b86fb24f15`. An independent SHA-256 comparison confirms that all 199 pinned files still match. The baseline remains intact. This is not a complete production audit pass. Application files remain untouched.
