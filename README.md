# Axcelerator

A VS Code extension that gives **GitHub Copilot (agent / chat tools)** the ability to drive **Microsoft Excel** end-to-end through the Language Model Tools API.

It uses [**xlwings**](https://www.xlwings.org/) as the engine:

- On **Windows**, xlwings talks to Excel through **COM** (`pywin32`), unlocking the full object model: `ListObjects` (Tables), `PivotTables`, `Queries` (Power Query / M), VBA macros, etc.
- On **macOS**, xlwings talks to Excel through **AppleScript**. Workbooks, sheets, ranges, formulas, tables, charts, Power Query, and VBA macros all work. (See [Microsoft's Mac VBA docs](https://support.microsoft.com/en-us/office/use-the-developer-tab-to-create-or-delete-a-macro-in-excel-for-mac-5bd3dfb9-39d7-496a-a812-1b5e8e81d96a). Some Windows-only constructs such as UserForms and ActiveX controls do not run on Mac.)

## What Copilot can do once installed

The extension registers 19 `vscode.lm` tools that Copilot can call autonomously in agent mode:

| Tool                                                                   | Purpose                                                                          |
| ---------------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| `excel_check_environment`                                              | Report OS, Excel version, what features are available                            |
| `excel_list_workbooks`                                                 | List open workbooks                                                              |
| `excel_open_workbook` / `excel_save_workbook` / `excel_close_workbook` | Workbook lifecycle                                                               |
| `excel_list_sheets` / `excel_add_sheet` / `excel_delete_sheet`         | Worksheets                                                                       |
| `excel_read_range` / `excel_write_range`                               | Read & write 2D values, with `expand` for auto-detected tables                   |
| `excel_set_formula`                                                    | Single cell, region, or dynamic-array formulas                                   |
| `excel_format_range`                                                   | Number format, bold/italic, font/background color, autofit                       |
| `excel_create_table`                                                   | Convert a range to an Excel **Table / ListObject**                               |
| `excel_create_chart`                                                   | Column / bar / line / pie / scatter / etc. anchored to a cell                    |
| `excel_create_pivot_table`                                             | Build a PivotTable from a Table (Windows: full COM; Mac: summarised table)       |
| `excel_add_power_query`                                                | Add or replace an M-expression query, optionally load to a sheet (Mac + Windows) |
| `excel_refresh`                                                        | Refresh a single query or all connections                                        |
| `excel_run_vba`                                                        | Run a named VBA macro (gated by setting; Windows + Mac)                          |
| `excel_run_python`                                                     | Run an arbitrary xlwings snippet against the live session (gated by setting)     |

This covers the use-cases:

- Workbook / sheet authoring
- Project plans + dashboards (tables, charts, formatting)
- Power Query (web, file, SharePoint, DB) -> Query Tables -> visuals
- Power-BI-style dashboards inside Excel
- Combining multiple SharePoint workbooks via Power Query
- DataFrames via xlwings (`xw.Range(...).options(pd.DataFrame).value`) -- exposed through `excel_run_python`
- Authoring formulas for end-user-friendly presentations

## Architecture

```
+--------------------+   stdin/stdout JSON-RPC   +----------------------+   xlwings   +----------+
| VS Code extension  | ----------------------->  | python/excel_bridge.py| --------->  |  Excel   |
| (TS, LM Tools API) | <-----------------------  |  (long-running)       | <---------  | (COM/AS) |
+--------------------+                           +----------------------+             +----------+
```

A single Python process is launched on first tool use and reused for the lifetime of the extension.

## Prerequisites

1. **Microsoft Excel** installed (Windows or Mac, Microsoft 365 recommended).
2. **Python 3.9+** with `xlwings`. Find your interpreter path with `which python3`, then install:

   ```bash
   python3 -m pip install xlwings
   # On Windows, also:
   python3 -m pip install pywin32
   ```

3. **VS Code 1.95+** with **GitHub Copilot Chat** enabled.

## Install the extension

**Step 1 -- build the `.vsix`** (run once from the `axcelerator` folder):

```bash
cd /path/to/axcelerator
npm install
npx @vscode/vsce@latest package --no-dependencies
```

**Step 2 -- install into VS Code**:

```bash
code --install-extension axcelerator-0.1.0.vsix
```

Reload VS Code (`Cmd+Shift+P` -> **Developer: Reload Window**). The extension is now active in every VS Code window -- no debug mode, no separate window required.

**To reinstall after any code change**, run both commands again:

```bash
npx @vscode/vsce@latest package --no-dependencies
code --install-extension axcelerator-0.1.0.vsix
```

## Configure

Open **Settings** (`Cmd+,`) and search **axcelerator**:

| Key                            | Default   | What it does                                                                                      |
| ------------------------------ | --------- | ------------------------------------------------------------------------------------------------- |
| `axcelerator.pythonPath`       | `python3` | Full path to the Python interpreter that has `xlwings` installed. Use `which python3` to find it. |
| `axcelerator.requestTimeoutMs` | `120000`  | Per-request timeout for Excel operations (ms)                                                     |
| `axcelerator.allowVba`         | `false`   | Lets Copilot call `excel_run_vba` (Windows + Mac, off by default for safety)                      |
| `axcelerator.allowPython`      | `false`   | Lets Copilot call `excel_run_python` (arbitrary xlwings code, off by default)                     |

After configuring, run **Axcelerator: Check Excel / xlwings Environment** from the Command Palette (`Cmd+Shift+P`) to confirm the bridge is working.

Every mutating tool call goes through the standard VS Code confirmation dialog before executing.

## Trying it from Copilot Chat

In agent mode, prompts like the following will cause Copilot to call the tools:

- "Open `~/Reports/Q3.xlsx`, add a sheet 'Summary', and put SUMIFS totals by region in A1:D10."
- "Create a Power Query that pulls data from a SharePoint list, clean it, and load it as a table on a new sheet."
- "Build a PivotTable from the `Sales` table grouped by Region (rows) and Quarter (columns) summing Amount, then add a clustered column chart next to it."

You can also reference the environment-check tool explicitly in chat with `#excelEnv`.

## Notes & Caveats

- The bridge requires Excel to be installed locally -- it does not talk to Excel Online.
- On macOS, the first call may prompt for **Automation** permission to control Excel. Allow it.
- `excel_run_python` exposes `xw`, `app`, and `wb` (when `workbook` is set). Assign to `result` to return a JSON-serialisable value.
- `excel_run_vba` and `excel_run_python` are disabled by default. Enable them in settings only when needed.

## Power Query on macOS

`excel_add_power_query` works on Mac by writing the **Mac-native DataMashup format** directly into the workbook (Excel for Mac has no scripting API for queries). The bridge:

1. Closes the workbook (Excel locks the file while open).
2. Patches the `.xlsx` ZIP with a `customXml/item1.xml` payload containing a UTF-16-wrapped, base64-encoded DataMashup binary that Excel for Mac reads.
3. Reopens the workbook (unless `noReopen=True`).

After injection the query is visible in **Data > Get Data > Launch Power Query Editor** under **Queries [N]**. The first time you open a freshly patched file Excel may show a yellow trust-bar -- click **Enable Content** to allow the query to refresh.

Limitations on Mac:

- `excel_refresh` is a no-op while the workbook is closed (the bridge surfaces a graceful message).
- `excel_run_vba` is supported on Mac as well as Windows. The workbook must be `.xlsm` with the macro defined, and macro execution must be permitted (Excel > Preferences > Security > Macro Security). UserForms / ActiveX controls remain Windows-only.
- Pivot tables created via the tool fall back to a static summarised table (no live PivotCache).

## Verification & testing

Run the full integration suite (60 tool invocations against a real Excel session) at any time:

```bash
cd /path/to/axcelerator
python3 test_all_tools.py
```

A green run prints `60/60 PASSED` and writes `TOOL_TEST_REPORT.md`. The script also leaves a sample workbook at `~/Developer/World Bank Mobile Phone Statistics - Axcelerator Test.xlsx` that exercises every tool, including a real Power Query (`Top10Query`) you can inspect in the Power Query Editor.

## Instructions for the agent (Copilot)

When using these tools autonomously:

- **Always call `excel_check_environment` first** to discover OS and Excel capabilities before planning multi-step work.
- **Workbook lifecycle**: prefer `excel_open_workbook` -> mutate -> `excel_save_workbook`. Use `excel_close_workbook` only when finished or before `excel_add_power_query` on Mac (the tool handles closing internally).
- **Power Query on Mac**: pass `loadToSheet`/`loadToCell` to make the query visible as a table on a sheet. The tool reopens the workbook automatically; do **not** call `excel_open_workbook` again immediately after.
- **Refreshing**: on Mac, queries refresh when the user clicks **Refresh** in the Power Query Editor or **Enable Content** on first open. `excel_refresh` is best-effort.
- **VBA / Python**: gated by user settings. If a call returns a "disabled" error, ask the user to enable the corresponding setting rather than retrying.
- **Idempotency**: `excel_add_power_query` with `replace=true` (default) overwrites an existing query of the same name without duplicating it.
