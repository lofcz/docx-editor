# Concurrent typing inside a proposed link

D30 and R19 remain open. This regression extends the tracked atom link checkpoint's shared fixture.

The original four cases edit outside the proposed link. The four added cases type inside its original source text. Both client orders and both review decisions use the same exact native and history oracles. The new cases fail before decisions, at the proposed-view assertion.

Input: `Before. Lead.  Tail. After.`

AI target: `Before. [Lead. <scio-mention key="user" value="Ada" /> Tail.](https://target.test "Keep") After.`

The offline participant inserts ` peer` at native offset 10, inside `Lead.`. Reconnection must retain that source boundary. The expected linked text starts with `Le peerad.`. Instead, the participant's text moves after the complete proposed link, before ` After.`. Both clients converge on that wrong placement.

The verified `f99195ee` runtime reproduces the failure through source and compiled packages:

- Source: four pass, four fail, 188 assertions. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-link-interior-baseline-source.log`.
- Compiled core/private SDK/evaluation transport: four pass, four fail, 188 assertions. Log: `/home/lofcz/.cache/sciobot-docx-native/d30-link-interior-baseline-compiled.log`.
- The original outside-range cases retain all history, exact decision, payload, geometry, and reopen checks. The added interior cases retain those same checks; the earlier placement failure prevents reaching them.

The native hyperlink structure writer creates both historical deletion and proposed insertion copies. Its struck-source cloning loses native source leaf identity and transaction-local slice recipes. That defect explains the lost source anchor, but preserving source identity alone does not solve revision membership for concurrent text.

A diagnostic source-identity experiment passed 51 native cases and the four original shared cases. The four interior cases then hid the participant's text in the deletion branch. This result rejects the experiment as a fix. The verified writer remains unchanged. The diagnostic patch is preserved at `/home/lofcz/.cache/sciobot-docx-native/d30-link-source-identity-experiment.patch`; its log is `d30-link-source-identity.log` in the same directory.

The next implementation must keep the participant's source boundary and causal revision membership through both native replacement branches. It must pass the existing proposed-view, exact accept/reject, both-actor history, decision history, and DOCX/room reopen oracles. Cloning source identities or relaxing the expected target does not satisfy that requirement.

Private types pass 2256 files. Scoped lint, max-lines caps, formatting, and whitespace checks pass. Compatibility policy passes 210 decisions with two cumulative format fields changed. This checkpoint adds no format field.

This checkpoint adds a failing regression and records a rejected experiment. It ships no runtime workaround and closes no TODO. The application remains read-only. Test-only transport uses the EigenPal Pro License.
