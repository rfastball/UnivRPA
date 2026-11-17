# Audit fixes (2026-11-17)

Scope: the 14 findings in the repository audit. Preserve pre-existing working-tree edits.

## Plan

- [x] 1/4: Open data workbooks with macros disabled; restore settings on every exit; match full paths.
- [x] 2/3: Refuse formatting populated non-script sheets; preserve trailing draft values and notes.
- [x] 5: Keep caller IF/loop state across callee control transfers and restore CALL frames.
- [x] 9/10/11: Reject invalid condition operators, reset brace search state, skip resolved ELSE_IF operands when already matched.
- [x] 6/7/8: Propagate foreground activation, clipboard and pixel-read failures.
- [x] 13: Build one program snapshot per batch edit.
- [x] 12/14: Document existing BBGGRR colors; remove unused SendKey wrappers.
- [x] Compile with vbaloop push, run existing and extended tests, inspect diff, release lease.

## Test delta brief

- Unique risk: wrong data source, unintended macros, formatter data loss, caller branch corruption, ignored API failures, and token/branch parsing failures not covered by the existing 21 tests.
- Existing owner: modTester.TC_ForeachRowExecution, TC_ScriptAssistFormatterRegression, TC_FlowNestedInteractions, TC_ParamGuardRuntime, TC_StaticVariableAnalysis, TC_AbortLatch and TC_FlowHappyPaths; clsActionTestDriver supplies deterministic external-call failures.
- Selected layer: array/state tests for control flow; existing temporary-sheet integration for editor behavior; injected OS/open failures for external boundaries. No live typing or clipboard overwrite in automated regressions.
- Expected cost delta: extend existing collected suites, no new per-test Excel launches, no real waits for new cases. Audit baseline: 21 pass, 0 fail, 0 skipped; command wall time about 6 seconds (not a benchmark).
- Evidence to retire: replace the formatter's single-row comment scenario with a batch scenario; remove unused keyboard wrappers. Keep existing input cleanup tests because they cover distinct held-key/button failures.

## Validation

- `vbaloop lint --json`: 41 source files checked, zero findings.
- `vbaloop push --json`: VBA compiled and UnivRPA.xlsm saved. Final source fingerprint: D1DB7F7799144EFD64C83F8E33F3D2158555B212E1B75A25B8C061B308E22617.
- Narrow owners: Test_FlowNestedInteractions and Test_StaticVariableAnalysis passed with zero failed assertions.
- Final `vbaloop test --json`: 21 passed, zero failed/skipped; no Excel crash or session loss.
- Inspected changes against the pre-task src backup, preserving earlier working-tree edits. No commit or distribution build.
- Existing color strings remain compatible; documents now use BBGGRR.

## Test delta

- Unique risk: wrong workbook selection or unexpected VBA execution, formatting data loss, lost caller IF/loop frames, eager resolution of skipped branches, invalid conditions, unmatched-brace analysis errors, and ignored external API failures.
- Owner and layer: seven existing TC_* suites in modTester; array/state checks, temporary worksheet integration, and deterministic external-call injection through the existing clsActionTestDriver.
- Added / extended / removed: extended seven suites and the existing driver; no new test files or public test entrypoints. Replaced single-row comment evidence with batch editing and strengthened the existing system-token analysis input with an unmatched suffix. Removed three unused production keyboard wrappers.
- Collected-case delta: 0 (21 before and after). Internal regression scenarios expanded within those existing owners.
- Runtime delta: not measured reliably; single full-command observations were about 6.0 seconds before and 6.6 seconds after, including Excel startup. This is not a benchmark.
- Heavy-resource launch delta: 0 per full suite; new cases use existing sheets/state and injected APIs, with no new live input or waits.
- Coverage-floor impact: no coverage policy or threshold changes; no numeric coverage measurement.

## Verification limits

- Actual Windows foreground policy, physical pixel reads, and live clipboard failure behavior were not exercised. Their success/failure paths were checked with the injected driver.
- Workbook-opening tests verify the security/event settings and restoration at the open boundary; no executable malicious workbook was used. AutomationSecurity ForceDisable protects VBA; it is not a general sandbox for external Excel files.
- Batch validation now shares one ProgramImage per edit batch; large-sheet latency was not benchmarked.
