# WinPC editor performance

Use `codex/react-block-proof` for progressive integration. Sciobot uses `claude/docx-native-cutover`.

## Current installed comparison

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
