# Retired paragraph collaboration metadata

R36 uses native canonical XML to retain an identity after a review decision removes its empty terminal paragraph. The retained identity lets later peer text arrive without reviving the rejected insertion. The evaluated transport uses its existing generic XML support. No Pro implementation changes occur.

The marker has this expanded element name:

```xml
<reviewState:retired
  xmlns:reviewState="urn:docx-editor.dev:review:retired-paragraph:1"/>
```

The marker occurs directly inside the paragraph's `w:pPr`. It has no attributes or children. Its prefix has no semantic meaning. Unknown names, future namespace versions, attributes, children, and identical elements inside opaque payloads stay opaque. Concurrent identical markers have the same idempotent meaning. The marker stores no old text, suggestion identity, or replacement content. Native Word revision addresses retain those responsibilities. Native formatting history can copy this marker into its old `w:pPr` snapshot. That historical copy does not retire the current paragraph.

Only an actor-bound native review transaction creates the marker. It acts when terminal paragraph mark resolution removes all inline content and a preceding paragraph remains. Solo decisions retain physical removal. Sections, required paragraphs after tables, opaque inline markers, and synchronized peer content retain their prior behavior. The new generic marker uses a fresh native node ID. An existing paragraph property container keeps its identity.

A marked paragraph remains in the live canonical store after its inline content disappears. Story traversal, layout, and binding omit it when another ordinary direct paragraph exists. Section marks and pending native property or paragraph-mark proposals keep the paragraph visible. If all direct paragraphs carry this empty state, the first stays available as the required paragraph. Any native inline content makes the retained paragraph visible. An independent peer suffix therefore retains its paragraph identity and native selection address. Explicit `deleteBlock` still removes the paragraph and wins over an offline edit beneath that deleted owner. Splits, joins, and fresh paragraph clones clear terminal retirement metadata when they author new paragraph structure. They preserve ordinary formatting and unrelated opaque metadata.

`writeOoxmlPackage` builds a pure Word export snapshot. It removes hidden empty retired paragraphs and recognized markers from surviving story paragraphs and native formatting history. It preserves required structure, peer content, opaque XML, and all unrelated parts. The live canonical package, raw collaboration state, and native undo history remain unchanged. Structural sharing and weak identity caches bound repeated projection work to changed subtrees.

Live collaboration persistence must preserve its existing raw room state. A DOCX export contains document content, not the room's retained identities or undo history. Pending native Word revisions still export normally. The metadata introduces no shared registry, journal, or transport schema fields. Older engines preserve it as generic XML but need an upgrade for its visibility and export behavior. The compatibility record is `.collaboration/changes/retired-paragraph-peer-recovery.json`.

## Resolved forward joins

R37 adds the empty `reviewState:joined` element in the same namespace. It has no attributes or children and occurs directly inside current `w:pPr`. Unknown namespaces, malformed marker shapes, and unrelated opaque payloads remain unchanged. A historical property snapshot does not set the current mark state.

An actor-bound forward review join retains both the source paragraph and its runs. The source carries `joined`; its existing native node IDs remain available to later peer operations. The marker means that the paragraph mark is resolved and absent in every review view. It does not create another pending revision or store replacement text. Empty sources stay outside ordinary story traversal when another required paragraph remains. A source with peer text joins the next paragraph under the same native owner. The survivor supplies paragraph properties; each span retains its original source paragraph and UTF-16 offset. A paragraph before a table joins into its first cell's first paragraph. That join does not continue through the table into a following paragraph. A source without a remaining forward target keeps its required final paragraph mark. Content-control and custom XML ownership remain join boundaries.

Layout applies this state in Original, Proposed, and All Markup. Live paragraph formatting projection and property rejection preserve current recovery metadata. Historical metadata cannot restore a removed recovery state. Weak identity caches reuse marker decisions for immutable paragraphs.

Ordinary splitting creates new paragraph marks before the source's original mark. Only the last resulting paragraph retains its resolved forward state. Ordinary joining inherits the second paragraph's final mark state. Fresh whole-paragraph copies clear both recovery marker kinds, including historical copies. Explicit owner deletion still removes that source, including concurrent text beneath the deleted owner. Undoing that deletion restores the source and retained peer text. Undoing original paragraph creation removes its owner; redo recovers the peer content with that owner.

Word export uses the same pure physical paragraph-join helper as solo native review decisions. It removes recognized recovery metadata and empty retired sources from its snapshot. It preserves content runs, unrelated XML, native pending content revisions, and unrelated parts. The live canonical package and collaboration history remain unchanged. Saving before unseen peer text arrives cannot consume or change its retained live identity. A DOCX remains a document snapshot rather than a serialized collaboration room.

The helper adds no shared operation, registry, codec, or transport fields. Older engines preserve generic XML but require an upgrade for the new visibility and export behavior. Compatibility decision: `.collaboration/changes/joined-paragraph-peer-recovery.json`.
