# WinPC editor performance

Use `codex/react-block-proof` for progressive integration. Sciobot uses `claude/docx-native-cutover`.

## Sciobot native block repairs

WinPC tested Sciobot adapter source `74b46313d` against engine `8f1564202`.
Remote text edits now retain unrelated block views. Folding releases the child view from its parent's slot.
The test also detects redundant style writes and checks that repaint restores three damaged properties.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Original adapter checks, 31/34 → 34/34 passed | — | Three failures resolved |
| Added content type check and ten block inset checks, 11/11 passed | — | 45/45 total, 1,050 assertions |
| Narrow toolbar and formatting checks, 28/28 passed | — | Includes wheel scrolling, selection, and undo |

The source type check passed across 2,076 files. The package build kept the engine external.
The toolbar repair is Sciobot commit `a9eb37343`.

Sciobot `b4d22c09f` ships the repaired adapter. WinPC verified all five archives and 766 installed files.
All 18 adapter outputs and 112 engine outputs matched their archives. Both builds resolved one engine.
Both participants passed reconnect, undo, redo, save, and reopen with 57 equations and zero duplicate widgets.

## Installed rich typing in Sciobot 24d9b82f0

Sciobot `24d9b82f0` ships engine `8f1564202`. WinPC verified five archives, 766 installed files, and the frozen lockfile.
The Home and collaboration builds each resolved one engine. All 77 runtime files matched the package.
WinPC alternated editor order across four pairs. Each editor typed the same 59 keys with 65 ms pauses.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Renderer CPU, 772.65 ms | 607.32 ms | 27.2% more CPU time |
| Elapsed including keyboard pauses, 5,305.15 ms | 5,244.00 ms | 1.2% more time |
| Median keypress to exact text, 7.00 ms | 1.10 ms | 5.90 ms longer |
| Median keypress to frame, 8.25 ms | 6.85 ms | 1.40 ms longer |
| 95th-percentile keypress to frame, 16.07 ms | 15.85 ms | 0.22 ms longer |
| Wrapping space to frame, 62.65 ms | 14.75 ms | 47.90 ms longer |
| Correct final text, caret, and backward selection, 4/4 | 4/4 | Equal |
| Follow-up input-method check, 4/4 | 1/4 | Three more passes |

The table uses medians across the four runs per editor. Both participant orders passed reconnect, undo, redo, save, and reopen. [All eight typing runs](profiles/winpc-installed-8f156-rich-typing.jsonl)

Chrome also sampled one 59-key run per editor on WinPC. Both finished with the expected text and zero browser errors.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Renderer CPU with profiling, 862.46 ms | 674.92 ms | 27.8% more CPU time |
| Browser selection writes, 50.36 ms sampled | 0.59 ms sampled | 49.77 ms more |
| Garbage collection, 48.79 ms sampled | 25.97 ms sampled | 22.82 ms more |

The DOCX scheduler used 419.89 ms of samples. Layout computation used 179.15 ms. Layout publication used 239.41 ms.
Within publication, painting used 66.18 ms, selection and caret updates used 64.55 ms, and post-paint observers used 55.14 ms.
These measurements include work performed by each function's children. [Paired CPU samples](profiles/winpc-rich-8f156-cpu-samples.json)

## Installed clipboard in Sciobot 24d9b82f0

WinPC alternated editor order across four pairs for each operation and document size.
Each elapsed time includes two animation frames after the keyboard command.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste 433 characters, 64.92 ms | 48.28 ms | 34.5% more time |
| Undo small paste, 41.07 ms | 32.29 ms | 27.2% more time |
| Redo small paste, 32.13 ms | 31.90 ms | 0.7% more time |
| Paste 158,633 characters, 144.04 ms | 109.65 ms | 31.4% more time |
| Undo large paste, 49.28 ms | 30.99 ms | 59.0% more time |
| Redo large paste, 62.06 ms | 54.22 ms | 14.5% more time |
| Small paste CPU, 33.90 ms | 22.60 ms | 50.0% more CPU time |
| Small undo CPU, 14.88 ms | 6.46 ms | 130.2% more CPU time |
| Small redo CPU, 12.97 ms | 5.86 ms | 121.3% more CPU time |
| Large paste CPU, 108.74 ms | 84.31 ms | 29.0% more CPU time |
| Large undo CPU, 21.03 ms | 6.75 ms | 211.5% more CPU time |
| Large redo CPU, 31.67 ms | 23.91 ms | 32.5% more CPU time |
| Exact final text, 24/24 | 24/24 | Equal |

Every run used the installed package and completed without a browser error.
All 33 consumer tests passed with 489 assertions. [All 48 clipboard runs](profiles/winpc-installed-8f156-clipboard.jsonl)

## Large paste undo and redo in Chrome

Three paired profiles used engine `5a0194e03` and Plate. Each restored the exact text after a 158,633-character paste.
The counts include every element in the browser document.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Elements before undo, 904 | 388 | 516 more |
| Elements after undo, 770 | 388 | 382 more |

The DOCX undo handler averaged 16.68 ms of CPU samples. Chrome spent 5.73 ms inside `Selection.setBaseAndExtent`.
The redo handler averaged 26.08 ms. The same selection operation used 7.41 ms.
These figures include only samples inside the history handler.

A separate inspection found that undo restored the exact original document part and root.
All three restored pages matched their previous records, but the engine rebuilt every record.
Redo rebuilt all 49 pages. All 49 matched their previous records. [Chrome profiles and page comparisons](profiles/winpc-history-attribution.zip)

### Reuse page records during undo and redo

WinPC compared engine `5a0194e03`, a layout checkpoint after every story edit, and Plate in all six orders.
Each of the 108 trials started with a fresh document. Elapsed times include two animation frames.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Undo 433 characters, 40.66 → 32.30 ms | 32.38 ms | Candidate took 0.3% less time |
| Undo 158,633 characters, 63.70 → 50.84 ms | 32.38 ms | Candidate took 57.0% more time |
| Redo 158,633 characters, 68.55 → 51.38 ms | 59.91 ms | Candidate took 14.2% less time |
| Large paste CPU, 108.04 → 112.16 ms | 84.39 ms | Candidate used 32.9% more CPU |
| Large undo CPU, 20.95 → 22.14 ms | 7.53 ms | Candidate used 194.1% more CPU |
| Large redo CPU, 31.71 → 23.90 ms | 23.51 ms | Candidate used 1.6% more CPU |

Large redo saved CPU in all six pairs. Large paste used more CPU in five pairs.
The candidate stayed out of the installed editor. [All 108 browser trials](profiles/winpc-history-layout-trials.jsonl)

The builds matched page, text, line, and selection geometry after paste, undo, and redo at both sizes.
All comparisons also passed after forced garbage collection. Both collaboration orders passed reconnect, undo, redo, save, and reopen. [Geometry counts and hashes](profiles/winpc-history-layout-geometry.json)

## Font and review reads during rich typing

WinPC compared engine `5a0194e03`, the font and review change, and Plate in all six orders.
Each editor typed 59 keys in the same paragraph. The control and candidate had equal final paint signatures.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Renderer CPU, 800.24 → 764.03 ms | 605.12 ms | Candidate used 26.3% more CPU time |
| Elapsed with keyboard pauses, 5,240.45 → 5,192.85 ms | 5,110.30 ms | Candidate took 1.6% more time |
| Median key to exact text, 7.10 → 7.05 ms | 1.10 ms | Candidate took 5.95 ms longer |
| Median key to frame, 8.25 → 8.20 ms | 6.85 ms | Candidate took 1.35 ms longer |
| 95th-percentile key to frame, 16.73 → 16.26 ms | 15.59 ms | Candidate took 0.67 ms longer |
| Wrapping key to frame, 58.85 → 53.35 ms | 10.55 ms | Candidate took 42.80 ms longer |
| Follow-up input-method check, 6/6 | 3/6 | Three more passes |

The median paired CPU saving was 39.22 ms. Five of six pairs used less CPU.
The 95th-percentile frame time fell in all six pairs. All 18 trials finished with the expected text, selection, and caret. [Browser measurements](profiles/winpc-rich-font-review.jsonl)

## Reuse after the edited rich paragraph

WinPC ran all six orders of engine `5a0194e03`, the changed line comparison, and Plate.
Each editor received 59 keys on the same rich lesson. The control and candidate had equal final paint signatures.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Renderer CPU, 766.41 → 762.31 ms | 623.46 ms | Candidate used 22.3% more CPU time |
| Elapsed with keyboard pauses, 5,266.05 → 5,239.75 ms | 5,002.65 ms | Candidate took 4.7% more time |
| Median key to exact text, 7.05 → 6.90 ms | 1.10 ms | Candidate took 5.80 ms longer |
| Median key to frame, 8.20 → 8.00 ms | 6.75 ms | Candidate took 1.25 ms longer |
| 95th-percentile key to frame, 15.57 → 16.19 ms | 15.89 ms | Candidate took 0.30 ms longer |
| Wrapping key to frame, 51.30 → 55.35 ms | 15.55 ms | Candidate took 39.80 ms longer |
| Exact final text, caret, and backward selection, 6/6 | 6/6 | Equal |
| Follow-up input-method check, 6/6 | 3/6 | Three more passes |

The candidate reused the remaining layout after all 58 edits that preserved line count. The wrapping edit ran the full path.
The median paired CPU saving was 36.76 ms. Four of six pairs used less CPU.
The 95th-percentile frame time increased in all six pairs. Two participants passed reconnect, undo, redo, save, and reopen in both orders. [All 18 browser runs](profiles/winpc-rich-paragraph-convergence.jsonl)

Across 60 frames, both builds matched 9,450 character rectangles, 3,395 controls, and 2,220 widget bounds.
The change stayed out of the installed editor because it increased frame time. [Geometry counts and matching hashes](profiles/winpc-rich-paragraph-geometry.json)

## Clipboard after the Sciobot upstream merge

Sciobot `2fd83fe0f` includes upstream `925da5ba2` and engine `5a0194e03`. WinPC ran four pairs per action and alternated editor order.
Each trial started with a fresh document. Times are medians from the keyboard command through two animation frames.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste 433 characters, 61.45 ms | 47.72 ms | 28.8% more time |
| Undo small paste, 40.54 ms | 32.58 ms | 24.4% more time |
| Redo small paste, 32.29 ms | 32.62 ms | 1.0% less time |
| Paste 158,633 characters, 143.21 ms | 118.67 ms | 20.7% more time |
| Undo large paste, 49.27 ms | 33.74 ms | 46.0% more time |
| Redo large paste, 65.66 ms | 48.79 ms | 34.6% more time |
| Large paste renderer CPU, 108.27 ms | 84.76 ms | 27.7% more CPU time |
| Large undo renderer CPU, 21.01 ms | 6.98 ms | 201.1% more CPU time |
| Large redo renderer CPU, 36.57 ms | 23.94 ms | 52.7% more CPU time |

All 48 documents had the expected text. Both production builds matched all 77 packaged engine files.
All 33 consumer tests passed with 489 assertions. Two peers passed undo, redo, reconnect, save, and reopen in both orders. [Browser measurements](profiles/winpc-merged-2fd83-history.jsonl)

## Browser work during 59 rich lesson keys

WinPC traced one run per editor on installed engine `5a0194e03`. Both editors received the same 59 keys.
An element counts once per style pass. These instrumented counts describe browser work during typing.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Elements styled across all passes, 3,387 | 266 | 12.73× as many |
| Elements styled in the largest pass, 2,859 | 141 | 20.28× as many |
| Browser layout passes, 61 | 61 | Equal |
| Animation frame callbacks, 718 | 72 | 9.97× as many |
| Time inside animation frame callbacks, 13.025 ms | 3.909 ms | 9.116 ms more |
| Time calculating styles, 60.781 ms | 8.655 ms | 52.126 ms more |

Eight toolbar subscriptions scheduled 472 of the DOCX callbacks.
Another trace showed 58 edits retained the same flow past the edited paragraph. One edit added a line.
Removing duplicate table sizing eliminated 236 scheduled style invalidations. Browser style time rose from 52.88 to 76.25 ms in that trial.
The installed comparison below provides repeated end-to-end timings. [WinPC trace counts](profiles/winpc-installed-5a019-rich-work.jsonl)

WinPC then tested one shared toolbar observer with the same controls. Each editor received 59 keys in all six orders.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Renderer CPU, control 785.28 ms → candidate 777.83 ms | 603.58 ms | Candidate used 28.9% more CPU time |
| Elapsed with keyboard pauses, 5,198.60 ms → 5,202.15 ms | 5,108.25 ms | Candidate took 1.8% longer |
| Median key to frame, 8.30 ms → 8.30 ms | 6.75 ms | Candidate took 1.55 ms longer |
| 95th-percentile key to frame, 16.13 ms → 16.57 ms | 15.61 ms | Candidate took 0.96 ms longer |

The candidate saved CPU in three pairs and improved 95th-percentile frame time in zero pairs.
The installed editor keeps its existing toolbar observers. [All 18 browser trials](profiles/winpc-toolbar-observer-trials.jsonl)

## Installed comparison at engine 5a0194e03

Sciobot commit `18edad0bb` pins engine `5a0194e03`. WinPC ran four alternating pairs for each action.
The times are medians through two animation frames. All 48 documents had the expected text.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste 433 characters, 64.12 ms | 47.77 ms | 34.2% more time |
| Undo small paste, 50.84 ms | 59.22 ms | 14.2% less time |
| Redo small paste, 32.35 ms | 31.67 ms | 2.1% more time |
| Paste 158,633 characters, 153.12 ms | 127.12 ms | 20.5% more time |
| Undo large paste, 51.29 ms | 62.03 ms | 17.3% less time |
| Redo large paste, 65.52 ms | 57.57 ms | 13.8% more time |
| Large paste renderer CPU, 109.99 ms | 84.13 ms | 30.7% more CPU time |
| Large undo renderer CPU, 21.36 ms | 7.44 ms | 187.1% more CPU time |
| Large redo renderer CPU, 33.69 ms | 24.63 ms | 36.8% more CPU time |

All 33 consumer tests passed with 489 assertions. WinPC verified five archives, 766 installed files, and the frozen lockfile.
Both builds resolved one engine. All 77 runtime files matched the archive. [Paired browser measurements](profiles/winpc-installed-5a019-history.jsonl)

Sciobot `2fd83fe0f` also pins this engine and includes upstream `925da5ba2`.
WinPC installed that complete source tree with its frozen lockfile and generated the locale modules.
All 289 frontend tests and 35 backend tests passed across the 52 changed test files.

Four alternating pairs typed 59 characters into the stored lesson, with a 65 ms keyboard delay.
Plate missed the target word in one additional attempt before timing began. The recorded repeat used the same word.
The percentiles use linear interpolation within each run. Each entry reports the median across four completed runs.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Typing renderer CPU, 762.59 ms | 609.90 ms | 25.0% more CPU time |
| Typing elapsed with pauses, 5,409.15 ms | 5,252.35 ms | 3.0% more time |
| Median keypress to frame, 8.20 ms | 6.85 ms | 1.35 ms longer |
| 95th-percentile keypress to frame, 16.08 ms | 15.97 ms | 0.11 ms longer |
| Selection placement before timing, 4/4 attempts | 4/5 attempts | One fewer failed attempt |
| Exact text, caret, and backward selection, 4/4 | 4/4 | Equal |
| Follow-up input-method check, 4/4 | 2/4 | Two more passes |

Browser style recalculation took 33.69 ms for DOCX and 7.63 ms for Plate. Browser layout took 7.20 ms and 7.84 ms.
These browser counters exclude the DOCX engine's JavaScript layout work. [All attempts](profiles/winpc-installed-5a019-rich-typing.jsonl)

## Previous installed comparison

Sciobot commit `69ae005f0` pins engine `267e76def`. WinPC tested that package with the typography fix from `8a92a3674`.
Each action used four alternating browser pairs. Times include two animation frames.
All 48 final documents matched the expected text. No trial reported a browser error.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste 433 characters, 71.39 ms | 47.93 ms | 48.9% more time |
| Undo small paste, 33.72 ms | 31.83 ms | 5.9% more time |
| Redo small paste, 32.03 ms | 32.18 ms | 0.5% less time |
| Paste 158,633 characters, 156.80 ms | 115.51 ms | 35.7% more time |
| Undo large paste, 60.35 ms | 42.86 ms | 40.8% more time |
| Redo large paste, 71.89 ms | 59.92 ms | 20.0% more time |
| Large paste renderer CPU, 123.97 ms | 84.42 ms | 46.8% more CPU time |
| Large undo renderer CPU, 24.61 ms | 7.61 ms | 223.4% more CPU time |
| Large redo renderer CPU, 46.56 ms | 23.77 ms | 95.9% more CPU time |

The installer checked all five archives, 766 installed files, and the lockfile. Both browser builds resolved one engine.
All 33 editing, typography, and collaboration tests passed with 489 assertions.
All 77 installed runtime files matched their package bytes. [Paired browser measurements](profiles/winpc-installed-267e-history.jsonl)

Four more pairs typed 59 characters into the same word of a stored lesson. The keyboard delay was 65 ms.
Each timing below is the median of four trial results. Each trial measured 59 keypresses.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Typing renderer CPU, 789.25 ms | 596.20 ms | 32.4% more CPU time |
| Typing elapsed with pauses, 5,463.40 ms | 5,451.65 ms | 0.2% more time |
| Median time from keypress to frame, 8.20 ms | 6.80 ms | 1.40 ms longer |
| 95th-percentile time from keypress to frame, 16.47 ms | 15.89 ms | 0.58 ms longer |
| Exact text, caret, and backward selection, 4/4 | 4/4 | Equal |
| Follow-up input-method check, 4/4 | 1/4 | 3 more passes |

The follow-up input-method check ran after timing. Plate timed out waiting for the expected composition text in three trials. [Timing and input checks](profiles/winpc-installed-267e-rich-typing.jsonl)

## Earlier installed comparison

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste 431 characters (four paired trials, median), 35.27 ms | 24.44 ms | 44.3% more time |
| Undo that paste (four paired trials, median), 18.20 ms | 8.02 ms | 126.9% more time |
| Redo that paste (four paired trials, median), 17.01 ms | 6.32 ms | 169.1% more time |
| Paste 158,633 characters through two frames (one trial per editor), 223.4 ms | 121.5 ms | 83.9% more time |

The HarfBuzz change saved 29.0 ms of layout and 28.2 ms of renderer CPU on the largest paste. Both values are median paired differences across four WinPC trials against the native baseline. The direct consolidated comparison appears below.

[Detailed Sciobot ledger and test conditions](https://github.com/sciocz/sciobot-next/blob/f16827cf8/docs/engineering/sciobot-port/evidence/winpc-editor-performance.md)

## Source history

The active source at consolidation includes upstream, shared history, caret preparation, and checked glyph widths. The archive merge preserves the other experimental commits with their original trees. An archived tip can contain an unsuccessful experiment, a superseded implementation, or an equivalent patch. Use `git show <commit>` to inspect a tip. Use `git diff a13a086e1...<commit>` to review its original divergence.

| Local branch at consolidation | Commit | Included in development history before archival |
| --- | --- | --- |
| `codex/browser-layout-performance` | `0601ad0fd73d` | Already integrated |
| `codex/react-block-proof` | `a13a086e1011` | Already integrated |
| `fix/experimental-peer-insert-suffix` | `0776de2d1208` | Already integrated |
| `fix/peer-insert-suffix` | `e13c3e68e1d5` | Archived |
| `fix/recorded-shared-redo` | `defa518ac3bc` | Already integrated |
| `fix/retained-review-joins` | `58ab9a16b051` | Archived |
| `integration/shared-history-upstream` | `18d4048fdc87` | Already integrated |
| `integration/upstream-hold-repair20` | `3b3361106b3f` | Already integrated |
| `main` | `f1175b8184b7` | Already integrated |
| `perf/bounded-prepared-opening` | `ef45a6578c70` | Archived |
| `perf/browser-caret-collapse` | `463354c869e5` | Archived |
| `perf/caret-after-selection` | `109380ff8128` | Archived |
| `perf/caret-endpoint-lookup` | `b145c4a4b7bd` | Archived |
| `perf/checked-shape-width` | `7768fb96c2d5` | Already integrated |
| `perf/combined-peer-picture-open` | `37bc9f95dafc` | Already integrated |
| `perf/control-boundary-delta` | `3977666583d0` | Archived |
| `perf/control-chrome-geometry` | `0fc52688658d` | Archived |
| `perf/deferred-projection-caret` | `4ed9cafe66f3` | Already integrated |
| `perf/first-disclosure-return` | `415ad276316b` | Archived |
| `perf/initial-caret-owner` | `2cd948ce5ce9` | Already integrated |
| `perf/initial-widget-reservations` | `0ca9035be542` | Already integrated |
| `perf/isolated-content-control-layers` | `80547b8e1e40` | Archived |
| `perf/latin-fitting-words` | `ba4635278309` | Archived |
| `perf/lazy-review-payloads` | `55725ba2d31b` | Archived |
| `perf/lesson-layout-resolution` | `180665d24cb2` | Archived |
| `perf/lesson-live-caret` | `8ca103c6489b` | Archived |
| `perf/lesson-native-word-ranges` | `5b3c58a91c36` | Archived |
| `perf/lesson-paced-publication` | `09acec39130e` | Archived |
| `perf/lesson-word-style-owner` | `59627138e44d` | Archived |
| `perf/paragraph-paint-publication` | `4ed9cafe66f3` | Already integrated |
| `perf/pin6-combined-caret` | `61dad648ab8e` | Archived |
| `perf/pin6-first-caret` | `373b4a8b7d46` | Archived |
| `perf/pin8-cell-widget-cumulative` | `d3b6e56611b8` | Archived |
| `perf/pin8-equation-layout` | `f543bc948ccb` | Archived |
| `perf/plain-probe-shaping` | `27982e6e0dc9` | Archived |
| `perf/plain-typing-publication` | `e8ad6c699c7c` | Already integrated |
| `perf/retain-text-tail` | `fcce7756d6f7` | Archived |
| `perf/reuse-shared-character-keys` | `d391678aba62` | Archived |
| `perf/reuse-xml-validation` | `cbf547bbb220` | Archived |
| `perf/same-page-translation-reviewed` | `26e804385718` | Archived |
| `perf/same-sheet-paragraph-translation` | `487ff6bff1a2` | Archived |
| `perf/scoped-inline-integration` | `23557b799855` | Archived |
| `perf/scoped-inline-widgets` | `177e23c93e0e` | Archived |
| `perf/scroll-wheel-ahead` | `b2e98fbd9c6d` | Archived |
| `perf/shared-control-properties` | `e770c078caa1` | Archived |
| `perf/shared-history-combined` | `7b49c3d380d8` | Already integrated |
| `perf/shared-paste-lookups` | `e7ceab007d14` | Already integrated |
| `perf/stable-text-lines` | `187cbaa527d8` | Archived |
| `perf/stable-virtual-paper` | `7c0afb787acf` | Archived |
| `perf/sync-projection-runtime` | `0abd4f3a9e55` | Archived |
| `perf/table-furniture-size-owner` | `1d497de458a3` | Archived |
| `perf/verified-rapid-typing` | `2043769497f5` | Already integrated |
| `perf/viewport-distance-evaluation` | `e911e292b874` | Archived |
| `perf/viewport-distance-overscan` | `f5c7c5f39fa7` | Archived |
| `perf/word-scoped-opening` | `4c31ccab340a` | Already integrated |
| `perf/word-widget-height-owner` | `515f0973deaa` | Already integrated |
| `winpc/combined-layout` | `421383d0956a` | Already integrated |
| `winpc/combined-layout-final` | `59c4188a4914` | Already integrated |
| `winpc/fitting-layout` | `d4a1c0a302c6` | Archived |

## HarfBuzz cold opening and typing

Four paired WinPC trials compared the installed native baseline with the packed width measurement change.
Positive values mean the candidate took longer. All 16 exact text and caret checks passed.

Ordinary documents changed by +9.65 ms in cold opening time, +17.27 ms in opening CPU, and −5.95 ms in typing CPU.
Stored lesson documents changed by +15.80 ms in cold opening time, +18.51 ms in opening CPU, and −1.60 ms in typing CPU.

Every stored lesson opening slowed. The experimental application keeps its previous package while we investigate.

WinPC ran 101 focused shaping and retained-selection tests against the consolidated source. All 565 assertions passed.

## Consolidated engine against Plate

WinPC built and packed the consolidated engine from `44468bde7`. The later attribution and ledger commits preserve its engine source.
Both editors used the same Sciobot browser build and pasted 158,633 characters. All eight final documents matched the expected text.
The four pairs alternated editor order. Times include two animation frames.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Pair 1, elapsed, 209.6 ms | 109.0 ms | 92.3% more time |
| Pair 1, renderer CPU, 168.5 ms | 84.9 ms | 98.5% more time |
| Pair 2, elapsed, 210.6 ms | 110.2 ms | 91.1% more time |
| Pair 2, renderer CPU, 172.1 ms | 85.7 ms | 100.8% more time |
| Pair 3, elapsed, 209.6 ms | 108.6 ms | 93.0% more time |
| Pair 3, renderer CPU, 170.7 ms | 81.9 ms | 108.4% more time |
| Pair 4, elapsed, 193.2 ms | 108.2 ms | 78.6% more time |
| Pair 4, renderer CPU, 168.8 ms | 84.5 ms | 99.8% more time |
| Median, elapsed, 209.6 ms | 108.8 ms | 92.6% more time |
| Median, renderer CPU, 169.7 ms | 84.7 ms | 100.4% more time |

### Large paste CPU profile

A separate Chromium sampling run used the same consolidated engine and the same 158,633-character paste.
The native layout call accounted for 99.0 ms of samples across three groups.

- Text width measurement inside paragraph breaking took 39.5 ms.
- Other paragraph breaking took 30.1 ms.
- Layout outside paragraph breaking took 29.4 ms.

The native selection write separately accounted for 15.2 ms of self samples.
Plate spent 19.0 ms of self samples in `toSlateRange`. These different functions do different work.
Profiler times come from one run and can vary from the elapsed medians above.

[Native CPU profile](profiles/consolidated-profile-4096-base-paste.cpuprofile.gz) and [Plate CPU profile](profiles/consolidated-profile-4096-plate-paste.cpuprofile.gz).

### Paste sensitivity to distinct numeric strings

Both inputs contained 4,096 sentences and 158,633 characters. The second input replaced each digit with zero.
Each input used four alternating editor pairs. All resulting documents matched their expected text.
The two inputs ran in separate batches.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Sequential numbers, elapsed, 209.6 ms | 108.8 ms | 92.6% more time |
| Sequential numbers, renderer CPU, 169.7 ms | 84.7 ms | 100.4% more time |
| Zeros of equal length, elapsed, 177.9 ms | 109.6 ms | 62.3% more time |
| Zeros of equal length, renderer CPU, 147.4 ms | 82.3 ms | 79.1% more time |

Repeated numeric text reduced median native CPU by 22.3 ms. The remaining elapsed gap measured 68.3 ms.
The sequential input made 4,683 width requests. Only 94 scalar requests repeated an earlier scalar key.
[Shaping counters and paired observations](https://github.com/sciocz/sciobot-next/blob/9af97f8f1/docs/engineering/sciobot-port/evidence/winpc-shaped-width.json)

## Opening 320 paragraphs with bounded early layout

Six WinPC trials per editor measured first paint, opening through the first exact edit, and first key latency.
Source candidate `ef45a6578` used smaller early layout steps than native control `f5c7c5f39`.
All 18 edit targets matched. Both native builds produced the same final geometry digest.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| First paint, 604.0 ms → 516.2 ms | 603.6 ms | 14.5% less time |
| Opening through first exact edit, 698.5 ms → 670.9 ms | 666.2 ms | 0.7% more time |
| First key latency, 30.1 ms → 42.7 ms | 29.9 ms | 42.8% more time |

The bounded candidate displayed text 87.4 ms before Plate. Its first key took 12.8 ms longer.
The consolidated source has different page preloading behavior. It requires a separate timing run.

## Paste across paragraph boundaries

The multiline input replaced 255 spaces with newline characters. Both inputs contain 158,633 characters.
Each row gives the median of four alternating editor pairs on WinPC. The two inputs ran in separate batches.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| One paragraph, elapsed, 209.6 ms | 108.8 ms | 92.6% more time |
| One paragraph, renderer CPU, 169.7 ms | 84.7 ms | 100.4% more time |
| 256 paragraphs, elapsed, 233.3 ms | 142.5 ms | 63.7% more time |
| 256 paragraphs, renderer CPU, 205.2 ms | 109.4 ms | 87.6% more time |

Native text matched the LF reference in all four multiline trials.
Plate readback contained 255 CRLF sequences. Normalizing those sequences to LF produced the exact expected text in every trial.

## Long paste after reducing the trailing-content scan

WinPC ran four alternating browser pairs against the packed engine. Both editors pasted 158,633 characters into the same document.
Elapsed time runs from the paste command through two animation frames. Each timing is the median for that editor.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste elapsed 200.8 ms | 109.1 ms | DOCX took 84.0% longer |
| Renderer CPU 171.8 ms | 83.4 ms | DOCX used 106.0% more CPU time |
| Exact text in 4 of 4 runs | 4 of 4 runs | Equal |

Four isolated measurements took 2.80–2.92 ms per scan before the change and 0.021–0.026 ms afterward.
All 4,000 comparisons produced identical suffix arrays. The previous browser batch used 169.7 ms of DOCX renderer CPU.

## Long paste with word fitting on each line

The paste produced one 159,118-character text run. The previous fitting loop stopped at offset 89.
The same document and input ran in four alternating browser pairs after word fitting resumed on later lines.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste elapsed 193.1 ms | 109.2 ms | DOCX took 76.9% longer |
| Renderer CPU 157.0 ms | 84.4 ms | DOCX used 86.0% more CPU time |
| Exact text in 4 of 4 runs | 4 of 4 runs | Equal |

The full page records matched the previous engine after paste, undo, and redo at all three tested input sizes.
The inputs contained 9,617, 38,825, and 158,633 characters. The largest produced 49 pages.
All 288 comparisons against the general line breaker matched across different widths, indents, kerning, and space compression rules.

## Paint the caret destination after a paste

WinPC ran four alternating browser pairs. Each editor pasted 158,633 characters into the same document.
The destination was page 47. DOCX built its visible pages once. The previous engine built them twice.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste elapsed 160.2 ms | 109.7 ms | DOCX took 46.0% longer |
| Renderer CPU 136.1 ms | 83.4 ms | DOCX used 63.2% more CPU time |
| Exact text in 4 of 4 runs | 4 of 4 runs | Equal |

The previous DOCX median was 193.1 ms. The caret position, line, scroll offset, and final page set matched after paste, undo, and redo.

## Paint words before a clipped space

Four alternating WinPC browser pairs pasted 158,633 characters. Both editors retained the exact text in all runs.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste elapsed 159.0 ms | 110.3 ms | DOCX took 44.2% longer |
| Renderer CPU 130.7 ms | 83.1 ms | DOCX used 57.2% more CPU time |

Across 81 painted lines, the native text spans fell from 486 to 108. All 154 shared browser hit tests gave the same model offsets.
Chrome compared 7,366 glyph rectangles. Their vertical coordinates and heights matched. The largest horizontal difference was 0.109375 CSS pixels.
The largest width difference was 0.015625 CSS pixels.

## Paste into 256 paragraphs after the viewport and paint changes

WinPC ran four alternating browser pairs at `87dcdfb44`. The input contained 158,633 characters and 255 paragraph breaks.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste elapsed 207.5 ms | 141.2 ms | DOCX took 47.0% longer |
| Renderer CPU 177.7 ms | 109.5 ms | DOCX used 62.3% more CPU time |
| Exact pasted text in 4 of 4 runs | 4 of 4 after CRLF normalization | Equal |

The earlier DOCX median was 233.3 ms. In a separate Chrome CPU profile, document layout used 125.4 ms.
Paragraph breaking used 76.8 ms within that interval. The edit transaction used another 16.1 ms.
Two collaborative editors also passed reconnect, undo, redo, save, and reopen in both peer orders.

## Typing across a line break in the lesson

Four alternating WinPC trials measured 59 keys in the React editor. Commit `dc9a7f260` retains paragraphs that only move vertically.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Renderer CPU 751.06 ms | 603.43 ms | DOCX used 24.5% more CPU time |
| Elapsed time including pauses 5,114.25 ms | 4,994.95 ms | DOCX took 2.4% longer |
| Frame containing the wrap 54.4 ms | 12.55 ms | DOCX took 333.5% longer |

The previous DOCX build used 799.79 ms of CPU and took 5,074.60 ms including pauses.
The wrap rebuilt 12 paragraphs, down from 42. All 696 geometry and accessibility records matched.
Both peer orders preserved 57 equations through reconnect, undo, redo, save, and reopen.

## Opening and navigating 320 paragraphs

WinPC ran six alternating trials per editor on the first paragraph. Three trials per editor targeted the last paragraph.
The tested DOCX opening change was `54257baba`. Commit `7bec4950d` integrates that change.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| First paint 517.55 ms | 594.55 ms | DOCX took 13.0% less time |
| Open through the first exact edit 636.55 ms | 664.60 ms | DOCX took 4.2% less time |
| First key 28.95 ms | 29.70 ms | DOCX took 2.5% less time |
| Wheel response 14.95 ms | 15.90 ms | DOCX took 6.0% less time |
| Open through the last-paragraph edit 745.10 ms | 663.70 ms | DOCX took 12.3% longer |

The native control took 708.40 ms through the first edit and 740.00 ms through the last edit.
Both native builds produced the same target coordinates and final document geometry.

### Merged opening, paint, and layout changes

The merged build at `de621fb98` used the same 320 paragraphs and the same browser input.
WinPC ran six trials per editor on the first paragraph and three on the last paragraph.
The native control used `44468bde7`.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| First visible text paint, 602.95 ms → 527.25 ms | 600.05 ms | 12.1% less time |
| First paragraph ready for a click, 592.85 ms → 515.85 ms | 574.95 ms | 10.3% less time |
| Open through the first exact edit, 710.10 ms → 640.85 ms | 670.80 ms | 4.5% less time |
| First key, 30.10 ms → 29.85 ms | 30.15 ms | 1.0% less time |
| First wheel response, 16.65 ms → 14.30 ms | 15.80 ms | 9.5% less time |
| Open through the last exact edit, 723.60 ms → 774.30 ms | 662.20 ms | 16.9% more time |
| Renderer CPU through the last edit, 600.48 ms → 799.60 ms | 557.84 ms | 43.3% more time |

Each trial placed the caret with one click and edited the expected character.
The two native builds produced identical rectangles for both clicked characters.
After each edit, all 67,298 character positions retained their text, source offset, page, vertical position, and height.
Across 27 pages, 2,640 horizontal positions differed by at most 0.125 CSS pixels.
Widths differed at 950 positions by at most 0.015625 CSS pixels.
The first page contained 2,339 character positions before editing. [Browser measurements and character audits](profiles/winpc-merged-opening-320.jsonl)

### Click the final paragraph while the document opens

Before `3fbf2c913`, the engine painted page 27 twice during opening in six Chrome traces.
In two traces, paragraph 318 moved under the pointer before the first click. A second click reached paragraph 320.
With the change, the engine painted page 27 once in all three traces.
Every first click reached paragraph 320 at the expected text offset of 149.
The first and last paragraph edits preserved all 67,298 character rectangles across 27 pages.
The initial 2,339 character rectangles also matched at 1/64 CSS pixel precision. [Chrome pointer events and character audits](profiles/winpc-opening-final-page.jsonl)

In both native builds, Ctrl+End moved from two visible pages to the final paragraph at offset 153.
The next key edited that paragraph. Undo restored its text.
An ordinary wheel scrolled 240 pixels while opening continued. The host handled Ctrl+wheel without scrolling.
Two connected editors opened with two visible pages each. Both delivery orders converged after concurrent edits, deletion, undo, and redo.
After reconnecting, both saved documents reopened with 319 paragraphs and the same digest.
Remote caret geometry matched on the first and final pages. [Keyboard, wheel, and collaboration checks](profiles/winpc-opening-final-navigation.jsonl)

WinPC ran four repository checks before opening, before the click fix, and after the click fix.
Workspace type checking passed at all three revisions after building the font declarations.
Each revision reported the same three formatting files, two fixture license headers, and collaboration release interval.
[Repository check results](profiles/winpc-opening-hook-comparison.jsonl)

WinPC ran all six orders of the native control, the closing fix, and Plate for each target.
The control used the merged opening build at `de621fb98`. The candidate included the extraction and `3fbf2c913`.
Each row gives the median of six trials per editor.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| First visible text paint, 511.50 ms → 514.80 ms | 609.30 ms | 15.5% less time |
| First paragraph ready for a click, 499.90 ms → 503.80 ms | 577.95 ms | 12.8% less time |
| Open through the first exact edit, 624.25 ms → 625.75 ms | 667.75 ms | 6.3% less time |
| Renderer CPU through the first edit, 652.89 ms → 552.89 ms | 539.65 ms | 2.5% more time |
| First key, 27.80 ms → 27.10 ms | 28.40 ms | 4.6% less time |
| First wheel response, 13.65 ms → 15.30 ms | 15.00 ms | 2.0% more time |
| Final paragraph ready for a click, 682.70 ms → 674.05 ms | 580.10 ms | 16.2% more time |
| Open through the final exact edit, 759.10 ms → 748.40 ms | 668.30 ms | 12.0% more time |
| Renderer CPU through the final edit, 668.37 ms → 659.74 ms | 552.37 ms | 19.4% more time |
| First key in the final paragraph, 29.85 ms → 29.90 ms | 29.80 ms | 0.3% more time |
| Wheel response from the final paragraph, 16.25 ms → 16.30 ms | 14.95 ms | 9.0% more time |

The control needed a second click in one final-paragraph trial. Its timing includes that retry.
The candidate and Plate reached both text offsets on the first click in all 12 trials each.
The native builds produced the same target rectangles and final geometry in every paired trial.
Three Chrome traces found 27.7–32.6 ms of React work before the first scroll event.
Browser layout then took 8.3–9.5 ms before pagination resumed. [Direct measurements and scroll traces](profiles/winpc-opening-final-interaction-320.jsonl)

### Closed ribbon tooltips during opening

The native ribbon mounted 29 tooltips, seven dropdown menus, and one font-size popover before Chrome delivered the first scroll event.
The installed `5a0194e03` build kept 32 empty tooltip portals in the document.
A bundle probe created a portal only while the tooltip appeared. The first hover created one portal, and closing removed it.
Pointer and keyboard tooltips kept their accessible names, anchor positions, and Escape behavior. The font menu opened and closed.

Six balanced rounds used the original bundle, the tooltip probe, and Plate.
All 18 trials reached the final word on the first click.
The two native builds produced identical target rectangles and document geometry. [Browser timings and tooltip checks](profiles/winpc-native-ribbon-portals.jsonl)

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Final paragraph ready for a click, 671.65 ms → 720.15 ms | 583.10 ms | 23.5% more time |
| Open through the final exact edit, 749.15 ms → 796.05 ms | 674.95 ms | 17.9% more time |
| Renderer CPU through the final edit, 701.69 ms → 788.33 ms | 611.68 ms | 28.9% more time |
| First key in the final paragraph, 30.20 ms → 29.40 ms | 29.35 ms | 0.2% more time |

The probe took longer through the final edit in four rounds. Sciobot retained the original tooltip.

## Memory allocation during multiline paste

Chrome sampled JavaScript allocation during one 158,633-character paste per editor at `87dcdfb44`.
The profiler included allocations that garbage collection reclaimed during the operation.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Estimated allocations 92.61 MB | 39.07 MB | DOCX allocated 2.37 times as much |

DOCX allocated an estimated 7.96 MB in text span creation, 3.24 MB copying spans for the break cache,
and 3.22 MB normalizing clipboard text.
The profiler recorded 11,157 DOCX samples and 4,539 Plate samples. Profiling times are excluded from the timing medians.

A separate one-paragraph read trace counted 1,838 placed lines containing 26,155 spans.
Readers requested the span array eight times on 1,641 lines before the second frame.
Page-field detection, paragraph lookup, line segments, and revision author collection each read spans outside the viewport.

## Shaping independent words together

Six WinPC trials compared the same paste on the DOCX control, the width change, and Plate.
Commit `267e76def` integrates the width change. The tested native control was `d875c5ece`.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste elapsed 184.2 ms | 115.0 ms | DOCX took 60.2% longer |
| Renderer CPU 150.6 ms | 83.6 ms | DOCX used 80.1% more CPU time |

Median paired elapsed time fell 15.0 ms. Five of six pairs improved. Median paired CPU fell 6.2 ms.
The native control medians were 192.8 ms elapsed and 158.4 ms CPU.
The change reduced HarfBuzz calls from 4,662 to 1,293. All nine paste, undo, and redo geometry checks matched.
Both peer orders preserved concurrent Unicode and numeric edits through reconnect, undo, redo, save, and reopen.
Ordinary opening measured 539.6 ms against 533.1 ms for the native control and 530.7 ms for Plate.

## Clipboard text normalization

WinPC compared `1021f9fb7` with the previous normalizer on 69,632 strings. Every result matched.
Across six trials per input, median normalization time fell from 1.676 to 0.072 ms for one paragraph.
It fell from 1.988 to 0.103 ms for multiline text and from 4.316 to 0.285 ms for Unicode text.
These measurements cover normalization alone. The 497 clipboard and paste tests passed in separate file processes.

## Installed Sciobot editor at `7ea4e425f`

Sciobot used its installed `87dcdfb44` engine. Both Home and the two-peer harness resolved one core package.
Each workload ran in four alternating editor pairs. Each operation ended after two animation frames.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Paste 433 characters 65.12 ms | 47.49 ms | DOCX took 37.1% longer |
| Undo the small paste 47.35 ms | 43.86 ms | DOCX took 8.0% longer |
| Redo the small paste 32.23 ms | 31.92 ms | DOCX took 1.0% longer |
| Paste 158,633 characters 160.01 ms | 136.40 ms | DOCX took 17.3% longer |
| Undo the large paste 50.34 ms | 32.02 ms | DOCX took 57.2% longer |
| Redo the large paste 76.17 ms | 57.60 ms | DOCX took 32.2% longer |
| Exact final text in 24 of 24 runs | 24 of 24 runs | Equal |

Large-paste renderer CPU measured 127.97 ms for DOCX and 83.35 ms for Plate.
Large-paste undo used 24.33 ms and 6.64 ms. Redo used 43.14 ms and 23.60 ms.
The timing batch included both input sizes and all three operations. Earlier paste-only batches used separate browser processes.
The raw browser measurements include elapsed time, renderer CPU, text length, exactness, and browser errors. [WinPC measurements](profiles/winpc-installed-87dc-history.jsonl)

## Rejected change to wrapped-word metrics and cached spans

Six WinPC triples tested every order of the native control, proposed change, and Plate.
The change combined run-local word measurements, clipboard normalization, and direct ownership of cached spans.
All 18 multiline pastes produced the expected text after line-ending normalization.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Control elapsed 192.65 ms | 141.95 ms | DOCX took 35.7% longer |
| Proposed change elapsed 201.16 ms | 141.95 ms | DOCX took 41.7% longer |
| Control CPU 168.03 ms | 109.40 ms | DOCX used 53.6% more CPU time |
| Proposed change CPU 170.35 ms | 109.40 ms | DOCX used 55.7% more CPU time |

The proposed change cost 4.43 ms elapsed and 2.41 ms CPU in the median paired comparison.
The branch retains the separately verified clipboard normalization change.

## Sciobot typography after reopen

Sciobot commit `8a92a3674` used installed engine `87dcdfb44`.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Fresh body text 16 px | 16 px | Equal |
| Fresh heading 30 px | 30 px | Equal |

The previous DOCX sizes were 18.4 px and 34.5 px.
All 33 consumer tests passed with 489 assertions. The three original test files retained their assertions.
The reload helper waited for a connected editor through the production load signal.
Both peer orders passed. Chrome selected Inter Regular Italic after reload and in a fresh browser context.

## Lesson overview with the installed editor

Four paired WinPC trials used Sciobot's lesson overview with installed engine `87dcdfb44`.
Each trial edited the same text at the same offset before opening and closing the overview.
The first two rows measure the initial response. The last two include the full animation.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Open overview, 165.45 ms | 40.15 ms | DOCX took 4.12 times as long |
| Close overview, 43.15 ms | 45.45 ms | DOCX took 5.1% less time |
| Open until motion stopped, 345.10 ms | 246.00 ms | DOCX took 40.3% longer |
| Close until motion stopped, 303.10 ms | 228.25 ms | DOCX took 32.8% longer |

During a separate Chrome trace, opening the overview moved the scroll position from 0 to 2,804 pixels.
The overview button kept focus while the editor scrolled to the previous text selection twice. [Paired measurements](profiles/winpc-overview-87dc-pairs.jsonl)

### Focused overview after repagination

Four WinPC trials compared the installed engine, the rebuilt engine with `9e796864e`, and Plate.
Each run edited the same paragraph before activating the overview button. The table shows median milliseconds.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Open overview, 150.10 ms → 90.65 ms | 41.95 ms | 116.1% more time |
| Close overview, 44.10 ms → 37.25 ms | 44.30 ms | 15.9% less time |
| Open until motion stopped, 280.40 ms → 245.15 ms | 257.05 ms | 4.6% less time |
| Close until motion stopped, 294.30 ms → 287.55 ms | 244.50 ms | 17.6% more time |

The fixed editor kept the button at 406 px and the scroll position at zero throughout opening.
It painted twice. The installed editor painted four times and scrolled 2,804 px. [Timing samples](profiles/winpc-overview-host-focus.jsonl)

The source regression reproduced a 1,771.7 px scroll before the fix. All 103 focused checks passed after the fix.
Both browser peer orders passed remote edits, reconnect, undo, redo, deletion, and save and reopen.
Eleven of 13 browser disclosure tests passed. Both DOCX builds failed the same wheel-completion and border-travel checks.
The border error measured 19.63 px in both DOCX builds against a 2 px limit.

### Overview motion with measured timeline geometry

Sciobot `2fd83fe0f` retained the running animations when Word replaced a paragraph.
It held the folded positions until the expanded layout included the measured timeline.
All 13 disclosure tests passed with 91 assertions. The reversed peer order passed 25 assertions.
Four WinPC trials used the same stored lesson and checked the edited paragraph. The table shows median milliseconds.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Open overview, 106.30 ms → 144.55 ms | 43.85 ms | 229.6% more time |
| Close overview, 45.40 ms → 45.60 ms | 38.25 ms | 19.2% more time |
| Open until motion stopped, 243.85 ms → 284.90 ms | 251.45 ms | 13.3% more time |
| Close until motion stopped, 295.55 ms → 295.25 ms | 229.95 ms | 28.4% more time |

The geometry fix added 38.25 ms to the initial opening response. [Timing samples](profiles/winpc-overview-2fd83-motion.jsonl)

## Move the measured overview timeline into its Word fragment

Sciobot commit `3be2ff6f1` measured the interactive timeline after activation. Word published seven pages once.
The same timeline control moved into the visible fragment at 668 × 303 px.
The heading stayed at y=406 px, and the scroll position stayed at zero.

Four balanced rounds compared `2fd83fe0f`, the previous transfer implementation, the live timeline, and Plate on the stored lesson.
The initial measurement includes two animation frames after the accessibility state changed. All 16 targeted edits matched.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Initial open, 107.00 ms → 84.55 ms | 40.85 ms | 107.0% more time |
| Initial close, 45.80 ms → 41.05 ms | 45.20 ms | 9.2% less time |
| Settled open, 281.55 ms → 274.45 ms | 257.65 ms | 6.5% more time |
| Settled close, 295.95 ms → 299.30 ms | 237.15 ms | 26.2% more time |

The median paired opening saved 14.45 ms initially and 10.85 ms through settlement.
Initial close saved 4.90 ms. Settled close added 4.00 ms in the paired comparison.
The four initial openings ranged from 41.20 to 140.90 ms.
Two additional observations put median cold paper movement at 165.85 ms before and 166.40 ms after.
The first visible heading moved from 21.90 to 26.65 ms. Cold settlement moved from 273.35 to 275.25 ms.
The second opening moved paper at 81.85 ms before and 52.25 ms after.

The first measurement implementation still rendered two timelines. Initial open rose from 118.55 to 122.25 ms.
Settled open rose from 290.20 to 290.90 ms. Initial and settled close added 0.35 and 0.75 ms.
Transferring the schedule data then raised initial open from 73.75 to 134.40 ms in its own four rounds.
Settled open improved by 19.50 ms. Initial and settled close added 0.40 and 1.65 ms.

WinPC passed 13 disclosure tests, 12 timeline tests, and both reversed peer cases with 221 assertions.
Reduced-motion changes during activation preserved focus and saved content. Consecutive activations of different controls committed both reading states.
Six rapid reversals per version ended folded with focus on the original trigger. [Timing samples and geometry](profiles/winpc-overview-3be2-timeline.jsonl)

One Chrome trace per editor followed the same lesson through opening and closing. The observers added marks without reading geometry.
Settled close took 295.3 ms for DOCX and 243.7 ms for Plate. The four-round unprofiled gap was 62.15 ms.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Close click to fold request, 246.2 ms | Click to final chevron transition, 227.0 ms | 19.2 ms longer |
| Fold request to motion release, 22.1 ms | No Word page publication | 22.1 ms additional |
| Motion release to settled observation, 27.0 ms | Final chevron transition to settled observation, 16.7 ms | 10.2 ms longer |

The host preserves its 220 ms closing motion. Plate uses a 200 ms chevron transition.
Word published five pages 20.07 ms after the fold request. Its motion state cleared 2.06 ms later.
The renderer then spent 12.94 ms updating styles and 2.24 ms on layout. Plate spent 0.06 ms updating styles.
During opening, Word published seven pages at 92.27 ms. Motion ended at 268.62 ms against 239.02 ms for Plate.
Opening settled at 301.58 ms against 256.63 ms for Plate. [Chrome timing marks and renderer intervals](profiles/winpc-overview-3be2-stages.json)

## Rejected first-return disclosure layout

The engine trial saved the preceding disclosure layout on its first activation. The first close reused all five pages.
The control placed 59 blocks. The trial placed zero. Both opened seven pages and placed 66 blocks.

Six balanced rounds compared the same `3be2ff6f1` host and Plate. All 18 targeted edits matched.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Initial open, 125.30 → 51.00 ms | 39.60 ms | 28.8% more time |
| Initial close, 44.05 → 33.40 ms | 41.60 ms | 19.7% less time |
| Settled open, 259.45 → 286.05 ms | 241.95 ms | 18.2% more time |
| Settled close, 294.15 → 285.30 ms | 228.75 ms | 24.7% more time |
| Entire settled cycle, 551.35 → 563.05 ms | 470.65 ms | 19.6% more time |

The median paired close saved 6.75 ms. Opening added 23.10 ms. The complete paired cycle added 5.70 ms.
Four preliminary rounds on the earlier host saved 14.05 ms per paired cycle. Settled opening rose from 273.65 to 274.05 ms.
Both exact-host origins used 115% zoom. All 908 text boxes, font properties, page boxes, headings, and timeline bounds matched.
The trial stayed outside the engine. [Paired measurements and geometry](profiles/winpc-overview-first-return.jsonl)

## Overview motion on the activation clock

Sciobot commit `5c5ad04c0` uses Plate's 200 ms overview timing. Each host window starts its heading and page motion at activation.
Six WinPC rounds exercised every order of the control, candidate, and Plate. Both native origins used installed engine `8f1564202`.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Initial open, 83.30 → 123.35 ms | 40.40 ms | 205.3% more time |
| Initial close, 43.10 → 44.80 ms | 44.55 ms | 0.6% more time |
| Settled open, 271.00 → 257.50 ms | 257.00 ms | 0.2% more time |
| Settled close, 293.10 → 261.75 ms | 243.75 ms | 7.4% more time |
| Entire settled cycle, 563.85 → 519.55 ms | 501.35 ms | 3.6% more time |

The paired cycle saved 42.60 ms. Five cycles and all six closes improved. Settled close saved 31.65 ms per paired trial.
Initial opening varied from 43.00 to 129.20 ms. Its paired median saved 0.05 ms, although its separate median rose 40.05 ms.
The remaining median cycle gap was 18.20 ms. The separate geometry check matched all 908 text boxes across the three views.
Both origins used 115% zoom and matched all pages, headings, and timeline bounds. [Six matched trials and geometry](profiles/winpc-overview-activation-clock.jsonl)

The original disclosure and timeline tests and both reversed peer cases passed with 221 assertions.
Reduced motion during activation and consecutive activations preserved focus and content. Type checks covered 217 files. Lint passed.

### First activation on the installed host

Eight openings on `5c5ad04c0` recorded frame order without CPU profiling. Chrome exposed the requested state within 4.1–5.3 ms.
Three openings reached the second frame before Word changed its page elements. Five reached it after Word and React committed.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Second frame before Word changed pages, 43.50 ms median | 41.60 ms median | 4.6% more time |
| Second frame after Word changed pages, 124.40 ms median | 41.60 ms median | 199.0% more time |

The later second frame coincided with earlier settlement, 250.50 ms against 276.00 ms. Both groups changed paper near 75 ms.
Six separate CPU traces captured the shorter path. The timeline read its dimensions during React's layout effect.
Chrome spent 14.19–21.05 ms in that callback. All renderer work between the final two frames took 45.72 ms median.
Plate used 2.27 ms in the same interval. [Frame order and CPU profiles](profiles/winpc-overview-first-activation.jsonl)

## Rejected line ownership summary

WinPC tested caching editable offsets and page fields as the layout placed each body line.
Six trials used every order of the control, candidate, and Plate on a 158,633-character single-paragraph paste.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Elapsed, control 160.52 ms → candidate 160.03 ms | 109.46 ms | Candidate took 46.2% more time |
| Renderer CPU, control 127.85 ms → candidate 127.94 ms | 83.30 ms | Candidate used 53.6% more CPU time |

The median paired change was 0.10 ms more elapsed time and 0.07 ms less CPU time.
All 18 documents matched the expected text. The candidate did not enter the engine.

## Paint the restored selection during undo and redo

Commit `e01c9376d` paints the selection after history restores it.
WinPC compared the change with engine `267e76def` on the same Sciobot build and with Plate.
Six orders per action produced 72 exact results. Times include two animation frames.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Undo 433 characters, 33.12 → 32.58 ms | 32.86 ms | 0.9% less time |
| Redo 433 characters, 40.52 → 44.23 ms | 41.59 ms | 6.3% more time |
| Undo 158,633 characters, 49.88 → 50.50 ms | 43.29 ms | 16.7% more time |
| Redo 158,633 characters, 66.69 → 50.81 ms | 73.26 ms | 30.6% less time |
| Large undo renderer CPU, 25.88 → 21.15 ms | 7.01 ms | 201.7% more CPU time |
| Large redo renderer CPU, 42.87 → 30.04 ms | 23.68 ms | 26.9% more CPU time |

The large redo saved 15.88 ms elapsed and 13.02 ms CPU in the median paired comparison.
The control failed the backward-selection paint test. The change passed it and all 52 focused tests. [Browser measurements](profiles/winpc-restored-history-selection.jsonl)
WinPC passed 128 additional tests with 489 assertions. Two peers converged in both editing orders after reconnect, undo, redo, and reopen.
The next character landed at the restored caret in all eight browser checks across DOCX and Plate.

## Wrapped words with different consumer zoom settings

Six WinPC trials compared engine `267e76def`, its wrapped-word change from `5a0194e03`, and Plate. The candidate painted all 137 text runs in the ordinary document at 115% of the control size. This difference prevents attribution of the timing changes to the engine.

| DOCX                                          | PLATE     | RELATIVE                           |
| --------------------------------------------- | --------- | ---------------------------------- |
| Single-paragraph paste, 157.45 ms → 142.11 ms | 109.15 ms | Candidate took 30.2% more time     |
| Single-paragraph CPU, 121.31 ms → 106.52 ms   | 83.11 ms  | Candidate used 28.2% more CPU time |
| Multiline paste, 191.86 ms → 192.26 ms        | 141.87 ms | Candidate took 35.5% more time     |
| Multiline CPU, 163.93 ms → 158.92 ms          | 108.69 ms | Candidate used 46.2% more CPU time |

All 18 engine geometry hashes matched across paste, undo, and redo at three sizes. The stored lesson matched all 390 painted text runs. Neither opening invoked the new wrap loop. Both collaboration orders matched all 224 word boxes after reconnect. Both saved and reopened the same Unicode text. [Browser measurements and geometry](profiles/winpc-wrapped-literal-comparison.jsonl)

### Wrapped words at matched zoom

WinPC rebuilt the same two engine packages against the combined Sciobot Home source.
Both builds matched all 527 opening text rectangles at 100% zoom in isolated browser contexts.
Six paste trials used every order of the control, candidate, and Plate.
Each paste contained 158,633 characters. Elapsed time includes two animation frames.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Single-paragraph paste, 149.90 ms → 133.87 ms | 109.52 ms | Candidate took 22.2% more time |
| Single-paragraph CPU, 120.14 ms → 106.18 ms | 83.35 ms | Candidate used 27.4% more CPU time |
| Paste 256 paragraphs, 200.00 ms → 192.46 ms | 141.58 ms | Candidate took 35.9% more time |
| Multiline CPU, 168.88 ms → 159.06 ms | 108.57 ms | Candidate used 46.5% more CPU time |
| Open the ordinary document, 550.55 ms → 548.95 ms | 509.05 ms | Candidate took 7.8% more time |
| Open the lesson, 917.25 ms → 919.65 ms | 905.05 ms | Candidate took 1.6% more time |

Single-paragraph paste saved 16.52 ms elapsed and 13.72 ms CPU in the median matched pair.
Multiline paste saved 0.97 ms elapsed and 8.55 ms CPU. Both formats used less CPU in all six pairs.
The single-paragraph paste made 944 HarfBuzz calls, down from 1,293.
All 18 geometry hashes matched after paste, undo, and redo. Neither opening invoked the wrapped-word loop. [Matched browser measurements](profiles/winpc-wrapped-words-matched.jsonl)

### Compact word boxes with eager body placement

WinPC compared installed engine `5a0194e03` and a candidate that stored compact word boxes until body placement.
Six trials used every order of both packages and Plate. Each paste contained 158,633 characters.

| DOCX | PLATE | RELATIVE |
| --- | --- | --- |
| Single-paragraph paste, 142.56 ms → 151.91 ms | 108.73 ms | Candidate took 39.7% more time |
| Single-paragraph CPU, 107.37 ms → 118.33 ms | 83.91 ms | Candidate used 41.0% more CPU time |
| Paste 256 paragraphs, 192.47 ms → 208.69 ms | 142.07 ms | Candidate took 46.9% more time |
| Multiline CPU, 161.54 ms → 173.62 ms | 107.66 ms | Candidate used 61.3% more CPU time |

The candidate added 9.34 ms elapsed and 5.10 ms CPU in the median single-paragraph pair.
Multiline paste added 15.37 ms elapsed and 11.59 ms CPU.
Three allocation profiles showed 57.35 MiB → 49.76 MiB. Sampled garbage collection changed from 7.48 ms to 7.17 ms.
Body placement rebuilt all word boxes and consumed 7.22 MiB there. The eager version stayed out of the release.

All 18 layout hashes and 527 opening text rectangles matched. Both collaboration orders matched 224 word boxes and reopened identical documents.
Separate read probes sampled every nineteenth line on this candidate and installed engine `8f1564202`.
All 93 sampled long-paragraph lines reached the page-field, ownership, and review-author checks. Paint read five. [Samples and reader stacks](profiles/winpc-compact-line-geometry.jsonl)
