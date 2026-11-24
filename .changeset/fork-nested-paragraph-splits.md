---
'@docx-editor.dev/pro': minor
---

Breaking collaboration upgrade (format `docx-collaboration:1.10.19.2`): concurrent paragraph splits keep text order inside a paragraph that an earlier concurrent split created, an insert of several lines made at the same time as a split keeps its line order, rejecting the tracked paragraph breaks of one participant keeps the breaks of the others, and undo and redo of such splits and inserts keep the text order, the text that others typed, and the other participants' breaks and formatting. A participant who joins a room shows the same break proposals as the others. Export saved rooms with your current release and reseed them, as described in [Upgrade fork rooms to format 1.10.19.2](https://www.docx-editor.dev/docs/2.x/pro/collaboration-versions#upgrade-fork-rooms-to-format-110192). Fixes #1972
