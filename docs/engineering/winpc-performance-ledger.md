# WinPC editor performance

Use `codex/react-block-proof` for progressive integration. Sciobot uses `claude/docx-native-cutover`.

## Measured results

| Installed native baseline operation         |   Native |    Plate | Sample                     |
| ------------------------------------------- | -------: | -------: | -------------------------- |
| Paste 431 characters                        | 35.27 ms | 24.44 ms | Four paired trials, median |
| Undo that paste                             | 18.20 ms |  8.02 ms | Four paired trials, median |
| Redo that paste                             | 17.01 ms |  6.32 ms | Four paired trials, median |
| Paste 158,633 characters through two frames | 223.4 ms | 121.5 ms | One trial per editor       |

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

| Document | Cold opening elapsed change | Cold opening CPU change | Typing CPU change |
| --- | ---: | ---: | ---: |
| Ordinary | +9.65 ms | +17.27 ms | −5.95 ms |
| Stored lesson | +15.80 ms | +18.51 ms | −1.60 ms |

Every stored lesson opening slowed. The experimental application keeps its previous package while we investigate.

WinPC ran 101 focused shaping and retained-selection tests against the consolidated source. All 565 assertions passed.

## Consolidated engine against Plate

WinPC built and packed the consolidated engine from `44468bde7`. The later attribution and ledger commits preserve its engine source.
Both editors used the same Sciobot browser build and pasted 158,633 characters. All eight final documents matched the expected text.
The four pairs alternated editor order. Times include two animation frames.

| Pair | Native elapsed, ms | Plate elapsed, ms | Native renderer CPU, ms | Plate renderer CPU, ms |
| ---: | ---: | ---: | ---: | ---: |
| 1 | 209.6 | 109.0 | 168.5 | 84.9 |
| 2 | 210.6 | 110.2 | 172.1 | 85.7 |
| 3 | 209.6 | 108.6 | 170.7 | 81.9 |
| 4 | 193.2 | 108.2 | 168.8 | 84.5 |
| Median | 209.6 | 108.8 | 169.7 | 84.7 |

### Large paste CPU profile

A separate Chromium sampling run used the same consolidated engine and the same 158,633-character paste.
The native layout call accounted for 99.0 ms of samples. The table divides this time into exclusive groups.

| Native layout work | Sampled time |
| --- | ---: |
| Text width measurement inside paragraph breaking | 39.5 ms |
| Other paragraph breaking | 30.1 ms |
| Layout outside paragraph breaking | 29.4 ms |

The native selection write separately accounted for 15.2 ms of self samples.
Plate spent 19.0 ms of self samples in `toSlateRange`. These different functions do different work.
Profiler times come from one run and can vary from the elapsed medians above.

[Native CPU profile](profiles/consolidated-profile-4096-base-paste.cpuprofile.gz) and [Plate CPU profile](profiles/consolidated-profile-4096-plate-paste.cpuprofile.gz).

### Paste sensitivity to distinct numeric strings

Both inputs contained 4,096 sentences and 158,633 characters. The second input replaced each digit with zero.
Each input used four alternating editor pairs. All resulting documents matched their expected text.
The two inputs ran in separate batches.

| Numeric text | Native elapsed | Plate elapsed | Native renderer CPU | Plate renderer CPU |
| --- | ---: | ---: | ---: | ---: |
| Sequential numbers | 209.6 ms | 108.8 ms | 169.7 ms | 84.7 ms |
| Zeros of equal length | 177.9 ms | 109.6 ms | 147.4 ms | 82.3 ms |

Repeated numeric text reduced median native CPU by 22.3 ms. The remaining elapsed gap measured 68.3 ms.
The sequential input made 4,683 width requests. Only 94 scalar requests repeated an earlier scalar key.
[Shaping counters and paired observations](https://github.com/sciocz/sciobot-next/blob/9af97f8f1/docs/engineering/sciobot-port/evidence/winpc-shaped-width.json)

## Opening 320 paragraphs with bounded early layout

Six WinPC trials per editor measured first paint, opening through the first exact edit, and first key latency.
Source candidate `ef45a6578` used smaller early layout steps than native control `f5c7c5f39`.
All 18 edit targets matched. Both native builds produced the same final geometry digest.

| Median | Native control | Bounded layout | Plate |
| --- | ---: | ---: | ---: |
| First paint | 604.0 ms | 516.2 ms | 603.6 ms |
| Opening through first exact edit | 698.5 ms | 670.9 ms | 666.2 ms |
| First key latency | 30.1 ms | 42.7 ms | 29.9 ms |

The bounded candidate displayed text 87.4 ms before Plate. Its first key took 12.8 ms longer.
The consolidated source has different page preloading behavior. It requires a separate timing run.

## Paste across paragraph boundaries

The multiline input replaced 255 spaces with newline characters. Both inputs contain 158,633 characters.
Each row gives the median of four alternating editor pairs on WinPC. The two inputs ran in separate batches.

| Paste input | Native elapsed | Plate elapsed | Native renderer CPU | Plate renderer CPU |
| --- | ---: | ---: | ---: | ---: |
| One paragraph | 209.6 ms | 108.8 ms | 169.7 ms | 84.7 ms |
| 256 paragraphs | 233.3 ms | 142.5 ms | 205.2 ms | 109.4 ms |

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
