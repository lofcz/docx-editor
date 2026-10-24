# Independent paragraph birth history

D30 and R19 remain open. This checkpoint fixes creation undo and redo for independent paragraphs containing peer text. Native rejection of that text remains a separate failing regression.

The transport previously retained table birth ancestry only. It now identifies independent paragraph births from the history step's own countable text items. Original relocation marks exclude split paragraphs. Restored text follows its original relocation chain. Foreign text that arrives later does not reclassify the creator's operation. The existing undo marker and ancestry retain only containers needed by surviving foreign content.

The application sequence normalizer previously removed retained native listings after the native undo filter kept them. It now leaves an existing retained listing under the pinned `docx-collaboration:1.6.5.1` wire profile. It never creates a missing listing. Cold updates declare existing root map constructors before inspecting the profile; they create no absent root or shared write. Ordinary string sequences retain their established removal behavior.

This changes deterministic interpretation of existing data, increasing repairVersion from 4 to 5. No new shared fields are added. The compatibility decision and export-and-reseed guide accompany the change. No production migration runs. The application adapter imports no Pro implementation. Evaluation transport sources use the EigenPal Pro License.

Verification:

- Four new real-session cases pass both client orders, tracked and direct paragraph births, creator undo/redo with exact pending native fingerprints, peer undo, creator undo without peer content, late peer redo, saved convergence, and room reopening. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-independent-birth-cold-final.log` (4 pass, 80 assertions).
- The existing split, moved-text undo, typing, generic sequence, and table birth checks pass 45 cases. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-paragraph-birth-transport-batch.log`. The four new fixture setup errors in that earlier run were corrected before the final run above.
- Format checks pass 19 cases, including refusal of the previous repair format without shared writes.
- The atom peer-history fixture initializes its empty payload store before room bootstrap, matching the shared-host contract. Its existing text, payload, native fingerprint, history, and saved assertions remain unchanged. Creation undo now preserves the peer paragraph; redo restores the exact native pending structure.

The combined final batch runs all eight selected files: independent paragraph history, generic sequences, split boundaries, moved-text undo, typing history, table birth history, format admission, and born atom shared history. It reports 72 pass and two native rejection failures across 74 cases, with 1712 assertions. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-independent-birth-integration-final.log`.

Pro types pass across 2315 files. Private types pass across 2251 files. The private build passes with 281 modules and native core external. Scoped lint, formatting, max-lines caps, and whitespace checks pass. Compatibility policy passes with 206 decisions. Logs use the `d30-independent-birth-` prefix under `/home/lofcz/.cache/sciobot-docx-native/`. Full C01/composer and compiled Pro/core verification remain pending. No baseline recapture, output normalization, or timeout increase occurs.

Remaining regression: after creation undo and redo pass, native rejection still removes the peer paragraph. The pending canonical tree places the peer's untracked text inside the AI insertion wrapper. That source wrapper therefore still claims the peer text. The next repair must give ordinary committed insertion its correct native ownership, while preserving imported Word structures and existing author-aware tracked insertion behavior. The two original failing cases remain intact.
