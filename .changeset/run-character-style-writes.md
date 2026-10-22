---
'@docx-editor.dev/core': minor
---

Add `characterStyle` to `setRunProperties`, so a run write can set or remove `w:rStyle` explicitly. Other run properties stay as authored, and tracked writes record the previous reference so rejecting restores it.
