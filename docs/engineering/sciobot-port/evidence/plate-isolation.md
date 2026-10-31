# Plate isolation of the native route

I21 evidence. Application branch `claude/docx-native-cutover`.

## Route graph

`WorkspacePanel.tsx` is a small entry. It renders `WorkspaceEditorSelector` from `docx/NativeWorkspaceRoute.tsx`. The Plate workspace (`LegacyWorkspacePanel.tsx`) loads through `lazy()` only when the legacy route renders: the native flag is off, or a conversion failed. The native editor (`docx/DocxWorkspacePanel.tsx`) also loads through `lazy()`.

Two native block views loaded Plate before this change:

- `blocks/code.tsx` imported the language list from `ui/code-block-node.tsx`, which pulls in the Plate block menus. The list now lives in `ui/code-block-languages.ts`.
- `blocks/media.tsx` imported `parseTwitterUrl` and `parseVideoUrl` from `@platejs/media`, which imports `platejs`. `src/lib/media-embed-url.ts` ports the two parsers. `media-embed-url.test.ts` compares both with the Plate parsers on 13 URLs.

## Bundle check

The check builds an entry with `bun build --target browser --splitting --metafile` and walks the static and dynamic imports from the entry. It reports every `platejs`, `@platejs/*`, and `slate` module that the walk reaches.

| Entry                                        | Modules | Plate packages |
| -------------------------------------------- | ------- | -------------- |
| `docx/DocxWorkspacePanel.tsx` before         | 1077    | 18             |
| `docx/DocxWorkspacePanel.tsx` after          | 730     | 0              |
| `WorkspacePanel.tsx` (entry), static imports | 188     | 0              |

The entry reaches Plate only through the dynamic import of `LegacyWorkspacePanel.tsx`.

## Retained consumers

[plate-consumers.json](plate-consumers.json) lists every application source file that imports Plate, with a category and a reason. The categories are:

| Category | Files | Reason |
| --- | --- | --- |
| `legacy-editor` | 217 | The Plate editor and its kits: the legacy workspace editor (flag off or failed conversion) and Markdown materials. |
| `legacy-elements` | 32 | Plate element renderers for lesson content in the legacy editor and Plate-based read views. |
| `cms-editor` | 12 | The CMS blog editor edits Markdown articles; Markdown formats keep Plate. |
| `slate-values` | 11 | Reads saved Plate values of legacy rows and print documents; native rows export their own DOCX bytes. |
| `plate-ai-command` | 6 | Server prompts for the Plate AI command menu of the legacy editor. |
| `legacy-review` | 5 | Plate suggestion and review helpers for the legacy editor. |
| `legacy-dialogs` | 2 | Dialog state for Plate equation and chemistry insertion. |
| `board` | 1 | The board lesson-flow widget renders Plate lesson values. |
| `material-transform` | 1 | Material transforms rewrite Plate values of legacy materials. |
| `shared-view` | 1 | Renders shared legacy lessons; native lessons render from their published projection. |
| `chat-bridge` | 1 | Chat tools for legacy and Markdown materials; native documents route to the on-demand native executors. |

Each category is outside the native route. The chat bridge (`useWorkspaceEditBridge.ts`) keeps a static `@platejs/markdown` import for legacy materials; native documents route to on-demand native executors.
