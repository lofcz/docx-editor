# Efficient implementation

Keep the requested TODO as the active implementation target until its completion criteria pass. Do not switch targets without user direction.

1. Read the TODO's completion criteria and inspect the current failure. Identify the smallest coherent implementation batch that addresses the cause.
2. Make the complete code batch before testing. Include related consumers, serialization, copying, and resolution paths when the change affects their shared invariant.
3. Start verification with the failing case, then the affected feature tests and type checks. Retain exact acceptance oracles and original performance limits.
4. Use full suites at a completed integration boundary, after focused failures are fixed. Do not repeat full suites after each small edit. Rerun a broad suite only when new code or unresolved evidence justifies it.
5. While slow verification runs, continue independent implementation or review in a separate worktree or isolated copy. Keep the tested source frozen and record which revision each result covers.
6. If repeated attempts do not move the completion criteria forward, reassess the implementation. Do not substitute more test permutations, status updates, or evidence prose for fixing the cause.
7. Run required repository checks. Close the TODO only when its actual criteria and deliverables are verified. Report completed TODOs separately from implementation progress and test counts.
8. Preserve unfinished work with a recoverable checkpoint. Keep progress notes concise and record only evidence that changes the next action or supports closure.

For R42, retain exact A01 acceptance output, exact rejection output, relationship isolation, save/reopen, history, and the required two-session behavior. Do not weaken those checks to gain speed.
