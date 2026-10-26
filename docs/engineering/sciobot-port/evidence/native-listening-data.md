# X09: native listening data and complete Word transcripts

Application commits `c5b7bbf8c` and `b5b8cee742ee9bb85cde4b36b9af685ca5643e77` implement the listening adapter in `src/lib/docx/listening.ts`. The adapter calls the existing shared grammar in `supabase/functions/_shared/activities/audio-exercise/markup.ts`. It uses the private native MDX codec for storage and source serialization. It imports no Plate parser, editor, or serializer. The existing B02/P05 payload and opaque-AST implementations remain the storage authority; a second fork-owned listening schema is unnecessary.

The payload retains its complete stored AST. Literal audio/hash/seed, stable key, speaker IDs, voices, portraits, roles, line text, unknown attributes, root metadata, nested extension data, and cached exercise fields survive document edits and export. Reads do not modify payloads. The keyless source fixture uses the persistent native block ID and invents no new audio identity. Known expression-backed fields and expression transcript lines refuse without evaluation. Existing literal braces, audio performance tags, ampersands, and apostrophes remain text.

`nativeListeningTranscript` generates the full title, introduction, and every named speaker turn. The application atomic fallback uses it for listening proposals. The native fixture represents transcript newlines with `w:br`; review stages retain the corresponding native manual breaks. No generated asset URL replaces a pristine reference. Audio hydration, playback, generation, and script/audio freshness reconciliation remain X10 and are not claimed here.

## Independent expected results

Both existing shared listening Markdown fixtures have explicit unit coverage: the full market dialogue and the keyless single-speaker block. The codec roundtrips through both the native AST serializer and the existing shared listening serializer. Parsed exercise data exactly matches `parseListeningMarkup` from the existing grammar. The transcript expectation is a complete independent literal string. Metadata read tests retain exact original JSON.

The browser fixture stores generated audio cache data and opaque fields alongside a complete native transcript. Nearby direct edits retain the entire payload and exact canonical native block fingerprint. Two AI transcript waves retain canonical baseline payloads during review. Both decisions compare complete payloads against independently modified original AST values, check the entire literal Word transcript, and leave nearby and peer text intact. Both author histories, remote caret geometry, saved convergence, and cold reopening pass.

Four additional offline cases delete a nearby paragraph while its peer edits it. Both client orders and both reconnect receivers retain exact payload and native listening fingerprints. Peer undo restores the independently edited paragraph; author undo/redo restores its edit. Saved convergence and cold reopening retain the listening data.

## Commands and results

- `bunx rstest run src/lib/docx/listening.test.ts src/hooks/chat/docxAtomicFallbacks.test.ts`: 12 pass; `/tmp/r19-listening-units-verified.log`.
- Installed atomic/listening browser batch: nine pass, 217 assertions; `/tmp/r19-listening-browser.log`.
- Final listening review file: four pass, 96 assertions; `/tmp/r19-listening-browser-verified.log`.
- Final offline listening subset: four pass, 44 assertions; `/tmp/r19-listening-offline.log`.
- Frontend types: 6,844 files; `/tmp/r19-listening-fe.log`. Final harness types: 2,418 files; `/tmp/r19-listening-harness-verified.log`. Scoped formatting, lint, and diff checks pass.

Runtime remains `4f0029f0b8df49f1aa966c43ce52a7a97b8443b0`, private archive `sciobot-docx-853c85888966.tgz`. No runtime/package refresh or collaboration format change was necessary. The evaluated transport remains test-only under the EigenPal Pro License. The product adapter uses the existing Apache core/private SDK and shared application grammar.

The production coverage snapshot needed only its application revision refreshed from `9a93e6b74435f760c4e3d2afa0aead47ce8e38bc` to `686f2b973e1ec1980e43e15953b342970337d93e`. All 199 source hashes, feature maps, kit membership, and print overrides are semantically identical. The snapshot refresh captures no new behavior and grants no additional feature closure.
