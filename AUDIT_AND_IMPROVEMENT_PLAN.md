# Axcelerator Audit and Improvement Plan

Date: 2026-04-29  
Scope: VS Code extension manifest, TypeScript extension host, Python xlwings bridge, packaging, documentation, and existing verification harness.

## Executive Summary

Axcelerator has a compact and promising architecture: 19 VS Code Language Model tools proxy to a long-running Python/xlwings bridge, and the TypeScript build currently passes. The existing integration report also shows a successful 60-call macOS Excel run.

The main risks are not basic build failures. They are feature-contract mismatches, platform-specific behavior gaps, packaging hygiene, and security boundaries around arbitrary Python/VBA execution. The highest-priority fixes are to align the LM tool schemas with implemented parameters, make Mac Power Query and PivotTable behavior explicit and reliable, harden bridge lifecycle/cancellation handling, and stop shipping test/generated artifacts in the VSIX.

## Verification Performed

| Check | Result | Notes |
| --- | --- | --- |
| TypeScript compile | Pass | `npm run compile` completed successfully. |
| VS Code diagnostics | Pass | No workspace errors reported. |
| Python syntax compile | Pass | `python -m py_compile python/excel_bridge.py test_all_tools.py` produced no errors. |
| npm dependency audit | Pass | `npm audit --json` reported 0 vulnerabilities. |
| Existing integration report | Pass, with caveats | `TOOL_TEST_REPORT.md` shows 60/60 passing on macOS/Excel/xlwings. Some tests validate graceful limitations rather than full feature support. |
| Existing VSIX inspection | Issues found | Package contains `test_all_tools.py`, `TOOL_TEST_REPORT.md`, and `python/__pycache__/excel_bridge.cpython-310.pyc`. |

## Issues Found

### Critical

| ID | Issue | Evidence | Impact | Recommended Fix |
| --- | --- | --- | --- | --- |
| C1 | Arbitrary Python execution has very broad capability once enabled. | `excel_run_python` executes user-provided code with `exec(...)` and exposes `xw`, `app`, and `wb`. The TypeScript gate is settings-based only. | A user or model with the setting enabled can run filesystem, network, process, or destructive Excel operations through Python. | Keep disabled by default, add workspace trust awareness, stronger confirmation text showing the code or a summary, optional allowlist mode, execution timeout/process isolation, and audit logging of snippet calls. |
| C2 | Mac Power Query load destination is advertised but not actually implemented as a load-to-sheet operation. | Schema exposes `loadToSheet`/`loadToCell`; Python Mac path passes them into `_pq_patch_xlsx`, but `_pq_patch_xlsx` ignores them and only injects the DataMashup. Returned text says refresh manually. | Copilot and users may believe data is loaded to a worksheet when only the query definition exists. This can produce incomplete workbook automation. | Either implement real load metadata/table wiring for Mac or change schema/docs to describe query-definition-only support. Add an integration assertion that the loaded table exists when `loadToSheet` is requested. |

### High

| ID | Issue | Evidence | Impact | Recommended Fix |
| --- | --- | --- | --- | --- |
| H1 | Implemented `noReopen` parameter is missing from the LM tool schema. | README and `test_all_tools.py` use `noReopen=True`; `package.json` schema for `excel_add_power_query` does not declare it. | Copilot cannot reliably discover or use the safer Mac Power Query flow that avoids Excel stripping injected DataMashup content after save. | Add `noReopen` to the schema with a clear description, or remove it from tests/docs if not intended as public API. |
| H2 | Mac PivotTable support is overstated. | `check_environment` reports `pivotTableSupported: true`; README says PivotTables work; implementation on Mac builds a static summarized table and ignores `columns` and `filters`. | Users may expect a real interactive PivotTable and get a static aggregation with only partial field support. | Report distinct capabilities such as `pivotTableInteractiveSupported` and `pivotSummarySupported`. Update docs and schema descriptions. Implement columns/filters or reject them on Mac with a clear error. |
| H3 | `excel_refresh` on Mac is best-effort but returns success even if refresh is unavailable or ignored. | Mac branch swallows exceptions from `book.api.refresh_all()` and returns `{"refreshed":"all"}` or the query name. | Copilot may proceed as if external data refreshed when nothing happened. | Return a capability-aware result with `attempted`, `supported`, and `warning` fields. If a specific `queryName` is requested on Mac and cannot be verified, return a warning or error. |
| H4 | Bridge cancellation and timeout do not stop in-flight Excel/Python work. | TypeScript `invoke` ignores the cancellation token; request timeout rejects the promise but does not cancel the Python operation. | A timed-out or cancelled tool call may keep mutating Excel and later produce unmatched responses. | Wire VS Code cancellation into request cancellation where possible. For long-running operations, serialize or track requests, terminate/restart the bridge after timeout, and clearly report possible partial completion. |
| H5 | Bridge startup error handling is incomplete. | `spawn` has no `error` handler. A bad Python path can leave startup waiting for the ready timeout and leave the process state unclear. | Users get slow or confusing failures when Python is missing or misconfigured. | Handle `proc.on('error')`, reject pending startup immediately, include configured interpreter path, and offer a command/action to open settings. |
| H6 | VSIX includes non-runtime artifacts. | Existing VSIX contains `test_all_tools.py`, `TOOL_TEST_REPORT.md`, and `python/__pycache__/excel_bridge.cpython-310.pyc`. | Package is larger, leaks local test assumptions, and ships generated bytecode. | Tighten `.vscodeignore` or add a `files` whitelist. Exclude tests, reports, `__pycache__`, `.pyc`, local workbooks, and generated VSIX files. |

### Medium

| ID | Issue | Evidence | Impact | Recommended Fix |
| --- | --- | --- | --- | --- |
| M1 | `~` paths are not expanded in the Python bridge. | `m_open_workbook` uses `os.path.abspath(path)` without `os.path.expanduser`; README examples use `~/Reports/Q3.xlsx`. | Natural Copilot/user prompts using `~` can fail or create/open the wrong path under the workspace. | Normalize all filesystem paths through a helper using `expanduser`, `expandvars`, and `abspath`. |
| M2 | Workbook lookup can select the wrong workbook when duplicate names are open. | `_find_workbook` accepts name, basename, or full path and returns the first basename match across all Excel apps. | Automation may target the wrong workbook if two files share a name. | Prefer exact full path when supplied, return an ambiguity error for duplicate basenames, and include candidate paths in the message. |
| M3 | Documentation and implementation disagree about VBA on Mac. | Header/README say VBA is Windows-only or unavailable; `m_run_vba` error text says Excel for Mac supports VBA with requirements. | Users get contradictory expectations about whether `excel_run_vba` should work on Mac. | Decide the support policy. If unsupported, fail immediately on Mac with concise text. If partially supported, update README, environment checks, and tests. |
| M4 | Mac Power Query XML patching uses regex/string edits on Open Packaging Convention XML. | `_pq_patch_xlsx` edits `[Content_Types].xml`, relationships, and workbook XML via regex. | Fragile against formatting, namespaces, duplicate relationships, and future Office changes. | Use XML parsers for relationships/content types, preserve namespaces, add package validation, and keep backup/restore tests around corrupt inputs. |
| M5 | Mac Power Query test preserves the DataMashup only by making injection the final operation. | `test_all_tools.py` comments state any xlwings save after injection strips the DataMashup package. | A normal workflow may accidentally destroy the query after insertion. | Make the tool return a prominent warning; consider marking workbook state as needing no further xlwings save or automatically reinjecting before final save. |
| M6 | Mac chart title path may silently fail. | Non-Windows branch attempts `chart.api.chart_title.set(...)` and suppresses errors. | Charts can be created without expected titles, but the tool reports success. | Verify title after setting or return a warning if setting it fails. Add chart metadata assertions to tests. |
| M7 | Configuration lacks validation ranges. | `requestTimeoutMs` is a number with no minimum or maximum. | Negative, zero, or tiny timeout values can cause confusing failures. | Add `minimum`, sensible `maximum`, and descriptions for long-running Excel operations. |
| M8 | Error responses to the model lose diagnostic detail. | Python returns tracebacks, TypeScript returns only `Error: ${msg}` to the LM result. | Debugging bridge failures requires opening the output channel and can slow issue resolution. | Log traces to the output channel and return a short error plus a correlation/request id. |
| M9 | TypeScript tool input types are untyped. | `LanguageModelTool<any>`, `summary?: (input: any)`, and `Record<string, any>` are used throughout. | Schema drift is easy, and missing fields are caught only at runtime. | Define TypeScript interfaces for each tool input and share/generate schema where possible. |
| M10 | Python bridge has limited input validation. | Most methods index directly into `p[...]` and rely on downstream xlwings errors. | Users receive low-context errors for missing/invalid fields; malformed ranges/colors can fail late. | Add per-method validation helpers and consistent error messages. |

### Low

| ID | Issue | Evidence | Impact | Recommended Fix |
| --- | --- | --- | --- | --- |
| L1 | Activation is broad. | Extension activates on `onStartupFinished` in every VS Code window. | Minor startup overhead and unnecessary registration in windows that never use Excel tools. | Prefer narrower activation events where supported, plus command activation for user commands. |
| L2 | Package metadata is incomplete for publication. | `repository.url` is `local`; package lacks a `license` field though a LICENSE file exists. | Marketplace/source consumers get poor metadata. | Set a real repository URL and license field. |
| L3 | Test harness depends on a local workbook path. | `test_all_tools.py` requires `~/Developer/World Bank Mobile Phone Statistics.xlsx`. | Tests are not reproducible on another machine or CI. | Generate fixture data or include a small sanitized fixture under test resources. |
| L4 | No automated npm test/lint pipeline exists. | `lint` script only echoes `no linter configured`; no `test` script. | Regressions can slip in outside manual Excel integration testing. | Add ESLint, unit tests for bridge lifecycle/schema mapping, and Python unit tests for pure helpers. |
| L5 | Existing report contains absolute local paths. | `TOOL_TEST_REPORT.md` includes `/Users/naveed/...`. | Reports are less portable and can leak local machine structure if packaged. | Use workspace-relative paths in generated reports or exclude reports from packages. |

## Improvement Opportunities

### Product and UX

| Improvement | Why It Helps | Suggested Approach |
| --- | --- | --- |
| Add a first-run setup/check flow. | Users need Python, xlwings, Excel permissions, and optional pywin32 on Windows. | Command palette wizard: detect interpreter, install guidance, run `check_environment`, open relevant settings. |
| Improve confirmation messages. | Mutating Excel actions can be destructive. | Include workbook, sheet/range, operation size, and for code execution show a collapsed preview or hash/request id. |
| Add dry-run/introspection tools. | Copilot can plan better before mutating workbooks. | Tools for workbook metadata, named tables, charts, queries, defined names, and workbook protection state. |
| Return structured warnings. | Current successful responses can hide degraded Mac behavior. | Standardize result shape: `ok`, `result`, `warnings`, `capabilities`, `partial`. |
| Add safer workbook targeting. | Reduces accidental edits. | Require full path for mutating operations when duplicate names are open; expose `workbookId` from `list_workbooks`. |

### Reliability

| Improvement | Why It Helps | Suggested Approach |
| --- | --- | --- |
| Add request queueing or explicit concurrency policy. | Excel automation is often stateful and not truly concurrent. | Serialize mutating calls and allow concurrent read calls only when proven safe. |
| Add bridge health checks. | Long-running Python/Excel automation can get wedged. | Ping method, process state tracking, automatic restart after timeout/exit, and visible output-channel diagnostics. |
| Make file operations transactional. | Power Query patching touches workbook ZIP internals. | Write to a new temp file, validate, then atomic replace. Keep backup only for rollback. |
| Add compatibility matrix. | Excel/xlwings behavior differs by OS and version. | Record tested versions, minimum supported Excel versions, and feature flags returned by `check_environment`. |

### Testing and CI

| Improvement | Why It Helps | Suggested Approach |
| --- | --- | --- |
| Split fast tests from Excel integration tests. | Most regressions should be caught without launching Excel. | Unit-test TS bridge lifecycle and Python pure helpers; keep Excel integration as optional/manual or tagged CI job. |
| Add schema contract tests. | Prevents drift like `noReopen`. | Compare `package.json` tool names/parameters to `TOOLS` and Python `METHODS`. |
| Add packaging tests. | Prevents accidental VSIX bloat/leaks. | Run `vsce ls` or inspect VSIX contents in CI and assert only expected files are present. |
| Add cross-platform test records. | Windows COM paths are not verified by the macOS report. | Maintain separate Windows and macOS reports, with platform-specific expected behavior. |

### Security and Governance

| Improvement | Why It Helps | Suggested Approach |
| --- | --- | --- |
| Add workspace trust integration. | VS Code extensions should respect untrusted workspaces before running code or touching files. | Disable `run_python`, `run_vba`, and possibly workbook file mutation in untrusted workspaces. |
| Add audit logging. | Makes destructive or code-execution actions reviewable. | Log timestamp, tool, workbook, sanitized params, and request id to the Axcelerator output channel. |
| Add path safety controls. | Prevents accidental writes outside intended locations. | Optional setting for allowed workbook roots and warnings for network/removable paths. |

## Recommended Implementation Roadmap

### Phase 1: Contract and Safety Fixes

1. Add `noReopen` to `excel_add_power_query` schema or remove it from public docs/tests.
2. Normalize filesystem paths with `expanduser`/`expandvars` in all path-accepting Python methods.
3. Add timeout validation in `package.json`.
4. Improve Python path startup errors with `spawn` error handling and actionable messages.
5. Update VBA-on-Mac documentation and runtime behavior so they agree.
6. Tighten `.vscodeignore` and rebuild the VSIX without tests, reports, bytecode, local workbooks, or prior VSIX files.

Acceptance criteria:

- `npm run compile` passes.
- `python -m py_compile python/excel_bridge.py test_all_tools.py` passes.
- VSIX contains only manifest, package/readme/license, `out/**`, and `python/excel_bridge.py`.
- README, schema, and integration test agree on all public parameters.

### Phase 2: Platform Capability Accuracy

1. Return granular capabilities from `check_environment`, especially for Mac PivotTable, Power Query load, refresh, and VBA.
2. Make Mac `excel_refresh` report warnings or unsupported status instead of unconditional success.
3. Make Mac PivotTable fallback reject unsupported `columns`/`filters` or implement them.
4. Add warning fields when chart title setting, refresh, or Power Query load cannot be verified.
5. Update README tables with exact Windows versus macOS behavior.

Acceptance criteria:

- Tool responses distinguish full support, fallback support, best-effort support, and unsupported features.
- Tests assert warnings/fallback notes, not just pass/fail.
- User-facing docs no longer overpromise Mac feature parity.

### Phase 3: Bridge Robustness

1. Wire cancellation tokens through `ExcelBridge.call`.
2. On timeout, mark bridge state uncertain and require or perform restart before the next call.
3. Add request ids to logs and returned errors.
4. Add a `health_check` or internal ping method.
5. Decide and enforce a concurrency model for mutating calls.

Acceptance criteria:

- Cancelling a VS Code tool call produces a clear cancellation result.
- Timed-out operations do not leave future calls using stale rejected startup promises.
- Output logs can correlate a tool call with bridge request/response/error events.

### Phase 4: Power Query Reliability

1. Replace regex XML/relationship edits with parser-based updates.
2. Write patched workbooks to a new temp file and atomically replace after validation.
3. Add package validation for required customXml parts and relationships.
4. Decide whether Mac load-to-sheet is supported; implement or remove from Mac-facing claims.
5. Add tests that save after injection and verify whether DataMashup survives or is intentionally reinjected.

Acceptance criteria:

- Corrupt or unusual workbook ZIP structures fail safely with rollback.
- Mac Power Query behavior is reproducible across at least two workbook fixtures.
- Documentation names the exact workflow required to preserve injected queries.

### Phase 5: Test and Release Maturity

1. Add ESLint and a real `npm test` script.
2. Add Python unit tests for pure helpers: path normalization, color parsing, DataMashup construction, and XML package patching.
3. Add schema contract tests across `package.json`, TypeScript `TOOLS`, and Python `METHODS`.
4. Add CI jobs for compile, lint, unit tests, npm audit, package inspection, and optional Excel integration.
5. Update package metadata with real repository and license fields.

Acceptance criteria:

- CI can validate the extension without a local private workbook.
- Manual Excel integration remains available but is no longer the only meaningful test.
- Release artifacts are reproducible and minimal.

## Suggested Priority Order

1. Fix schema/docs drift: `noReopen`, Mac Power Query load claims, Mac VBA wording.
2. Fix packaging hygiene and release metadata.
3. Harden bridge startup, timeout, and cancellation behavior.
4. Clarify and test Mac fallback capabilities for PivotTable, refresh, chart titles, and Power Query.
5. Build fast automated tests and CI around schema contracts and package contents.
6. Revisit arbitrary Python/VBA execution with workspace trust, richer confirmations, and audit logs.

## Open Questions

1. Should `excel_run_vba` be supported on macOS as a best-effort feature, or intentionally Windows-only?
2. Should Mac Power Query aim for true load-to-sheet support, or only query-definition injection?
3. Should workbook mutation require full path or a workbook id instead of accepting ambiguous names?
4. Should arbitrary Python execution remain a first-class tool, or move behind an even stronger experimental setting?
5. Should the integration test generate its own fixture data so it can run on CI and other machines?