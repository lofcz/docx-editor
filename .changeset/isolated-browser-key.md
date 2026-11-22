---
'@docx-editor.dev/core': patch
---

The editor applies an isolated keystroke before the next timer when the browser reports an empty input queue. Queued input keeps its existing batching behavior.
