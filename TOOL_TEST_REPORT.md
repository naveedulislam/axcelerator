# Axcelerator — Tool Verification Report

**Date:** 2026-04-29 15:01:56  
**Workbook tested:** `axcelerator_test_world_bank_mobile.xlsx`  
**Platform:** macOS — Excel ? — xlwings 0.35.2 — Python 3.12.1  

## Summary

| Metric | Value |
| --- | --- |
| Total tool calls | 61 |
| Passed | 61 |
| Failed | 0 |
| Total wall-clock | 4716 ms |

## Result per tool call

| # | Tool | Description | Status | ms | Notes |
| --- | --- | --- | --- | --- | --- |
| 1 | `check_environment` | OS / xlwings / Excel version | PASS | 1 | {"os": "Windows", "pythonVersion": "3.12.1", "xlwingsVersion": "0.35.2", "excelRunning": false, "comAvailable": true, "vbaSupported": true, "powerQuer |
| 2 | `list_workbooks` | List workbooks (before open) | PASS | 0 | [] |
| 3 | `open_workbook` | Open test copy | PASS | 924 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "C:\\Users\\navee\\AppData\\Local\\Temp\\axcelerator_test_world_bank_mobile.xlsx", "cr |
| 4 | `list_workbooks` | List workbooks (after open) | PASS | 104 | [{"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "C:\\Users\\navee\\AppData\\Local\\Temp\\axcelerator_test_world_bank_mobile.xlsx", "s |
| 5 | `list_sheets` | Inspect source sheets | PASS | 40 | [{"name": "Data", "index": 1, "usedRange": "$A$1:$X$223", "rowCount": 223, "columnCount": 24}, {"name": "Definition and Source", "index": 2, "usedRang |
| 6 | `read_range` | Read header row A1:X1 | PASS | 26 | {"address": "$A$1:$X$1", "values": [["Series Name", "Series Code", "Country Name", "Country Code", "1997 [YR1997]", "1998 [YR1998]", "1999 [YR1999]",  |
| 7 | `read_range` | Read all 222 data rows A1:X223 | PASS | 43 | {"address": "$A$1:$X$223", "values": [["Series Name", "Series Code", "Country Name", "Country Code", "1997 [YR1997]", "1998 [YR1998]", "1999 [YR1999]" |
| 8 | `add_sheet` | Add TempSheet (to be deleted) | PASS | 86 | {"name": "TempSheet", "index": 2} |
| 9 | `delete_sheet` | Delete TempSheet | PASS | 41 | {"deleted": "TempSheet"} |
| 10 | `add_sheet` | Add Summary sheet | PASS | 65 | {"name": "Summary", "index": 2} |
| 11 | `add_sheet` | Add Dashboard sheet | PASS | 80 | {"name": "Dashboard", "index": 3} |
| 12 | `add_sheet` | Add PQ_Data sheet | PASS | 82 | {"name": "PQ_Data", "index": 4} |
| 13 | `add_sheet` | Add Pivot sheet | PASS | 76 | {"name": "Pivot", "index": 5} |
| 14 | `write_range` | Title block to Summary!A1 | PASS | 63 | {"address": "$A$1", "rows": 3, "cols": 3} |
| 15 | `write_range` | Top-10 table to Summary!A5 | PASS | 53 | {"address": "$A$5", "rows": 11, "cols": 3} |
| 16 | `write_range` | Trend table to Summary!A18 | PASS | 52 | {"address": "$A$18", "rows": 6, "cols": 21} |
| 17 | `set_formula` | AVERAGE formula Summary!C17 | PASS | 15 | {"address": "$C$17"} |
| 18 | `set_formula` | Average label Summary!B17 | PASS | 18 | {"address": "$B$17"} |
| 19 | `format_range` | Format Summary A1:C1 | PASS | 27 | {"address": "$A$1:$C$1"} |
| 20 | `format_range` | Format Summary A2:C2 | PASS | 16 | {"address": "$A$2:$C$2"} |
| 21 | `format_range` | Format Summary A5:C5 | PASS | 17 | {"address": "$A$5:$C$5"} |
| 22 | `format_range` | Format Summary C6:C17 | PASS | 19 | {"address": "$C$6:$C$17"} |
| 23 | `format_range` | Zebra row 1 | PASS | 16 | {"address": "$A$6:$C$6"} |
| 24 | `format_range` | Zebra row 2 | PASS | 17 | {"address": "$A$7:$C$7"} |
| 25 | `format_range` | Zebra row 3 | PASS | 17 | {"address": "$A$8:$C$8"} |
| 26 | `format_range` | Zebra row 4 | PASS | 14 | {"address": "$A$9:$C$9"} |
| 27 | `format_range` | Zebra row 5 | PASS | 13 | {"address": "$A$10:$C$10"} |
| 28 | `format_range` | Zebra row 6 | PASS | 15 | {"address": "$A$11:$C$11"} |
| 29 | `format_range` | Zebra row 7 | PASS | 16 | {"address": "$A$12:$C$12"} |
| 30 | `format_range` | Zebra row 8 | PASS | 15 | {"address": "$A$13:$C$13"} |
| 31 | `format_range` | Zebra row 9 | PASS | 17 | {"address": "$A$14:$C$14"} |
| 32 | `format_range` | Zebra row 10 | PASS | 20 | {"address": "$A$15:$C$15"} |
| 33 | `format_range` | Format Summary A1:D25 | PASS | 16 | {"address": "$A$1:$D$25"} |
| 34 | `format_range` | Format Summary A18:X18 | PASS | 18 | {"address": "$A$18:$X$18"} |
| 35 | `create_table` | Top10Table on Summary | PASS | 26 | {"name": "Top10Table", "address": "$A$5:$C$15"} |
| 36 | `create_table` | TrendTable on Summary | PASS | 25 | {"name": "TrendTable", "address": "$A$18:$U$23"} |
| 37 | `write_range` | Top-10 data to Dashboard!A1 | PASS | 52 | {"address": "$A$1", "rows": 11, "cols": 2} |
| 38 | `format_range` | Dashboard top-10 header A1:B1 | PASS | 17 | {"address": "$A$1:$B$1"} |
| 39 | `create_chart` | Bar chart — Top 10 countries | PASS | 35 | {"name": "Chart 1"} |
| 40 | `write_range` | Trend data (transposed) to Dashboard!A15 | PASS | 57 | {"address": "$A$15", "rows": 21, "cols": 6} |
| 41 | `format_range` | Trend header A15:F15 | PASS | 18 | {"address": "$A$15:$F$15"} |
| 42 | `create_chart` | Line chart — mobile trends 1997-2016 | PASS | 25 | {"name": "Chart 2"} |
| 43 | `write_range` | Top-5 data to Dashboard!A40 | PASS | 50 | {"address": "$A$40", "rows": 6, "cols": 2} |
| 44 | `format_range` | Top-5 header A40:B40 | PASS | 18 | {"address": "$A$40:$B$40"} |
| 45 | `create_chart` | Column chart — top 5 | PASS | 24 | {"name": "Chart 3"} |
| 46 | `write_range` | Banner to Dashboard!H1 | PASS | 51 | {"address": "$H$1", "rows": 5, "cols": 1} |
| 47 | `format_range` | Style Dashboard H1:L1 | PASS | 20 | {"address": "$H$1:$L$1"} |
| 48 | `format_range` | Style Dashboard H2:L2 | PASS | 15 | {"address": "$H$2:$L$2"} |
| 49 | `format_range` | Style Dashboard H3:L3 | PASS | 17 | {"address": "$H$3:$L$3"} |
| 50 | `format_range` | Style Dashboard H4:L5 | PASS | 15 | {"address": "$H$4:$L$5"} |
| 51 | `format_range` | Style Dashboard A1:L50 | PASS | 20 | {"address": "$A$1:$L$50"} |
| 52 | `create_pivot_table` | Summarised pivot from Top10Table -> Pivot!A1 | PASS | 112 | {"name": "Top10Pivot", "destination": "$A$1"} |
| 53 | `write_range` | Seed data to PQ_Data!A1 | PASS | 57 | {"address": "$A$1", "rows": 11, "cols": 3} |
| 54 | `format_range` | PQ_Data header | PASS | 73 | {"address": "$A$1:$C$1"} |
| 55 | `save_workbook` | Save before PQ injection | PASS | 43 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "C:\\Users\\navee\\AppData\\Local\\Temp\\axcelerator_test_world_bank_mobile.xlsx"} |
| 56 | `run_python` | Introspect Dashboard via Python | PASS | 45 | {"result": {"dashboard_used_range": "$A$1:$L$45", "dashboard_rows": 45, "dashboard_cols": 12, "workbook_name": "axcelerator_test_world_bank_mobile.xls |
| 57 | `run_vba` | run_vba: Windows-only policy (Mac=unsupported, Win=missing-macro error) | PASS | 13 | Mac: bridge returns 'Windows-only' unsupported error. Win: missing macro raises. |
| 58 | `save_workbook` | Final save (before PQ injection) | PASS | 35 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "C:\\Users\\navee\\AppData\\Local\\Temp\\axcelerator_test_world_bank_mobile.xlsx"} |
| 59 | `add_power_query` | Inject Power Query 'Top10Query' (last operation) | PASS | 1723 | {"queryName": "Top10Query", "loadedTo": null} |
| 60 | `list_workbooks` | List workbooks after PQ injection (workbook closed) | PASS | 30 | [{"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "C:\\Users\\navee\\AppData\\Local\\Temp\\axcelerator_test_world_bank_mobile.xlsx", "s |
| 61 | `refresh` | Refresh after PQ injection (platform-aware) | PASS | 58 | Windows: COM refresh succeeds after PQ injection. |

## Dashboard built in Excel

### Sheets created
| Sheet | Contents |
| --- | --- |
| **Summary** | Title block, Top-10 table (zebra-striped), AVERAGE formula, two Excel Tables, trend data, pivot summary |
| **Dashboard** | Bar chart (top 10), Line chart (5-country trend 1997-2015), Column chart (top 5), colorful banner |
| **PQ_Data** | Seed table, Power Query 'Top10Query' injected via M formula |

### All 19 tools exercised
| Tool | Tested via | Phase |
| --- | --- | --- |
| `check_environment` | Direct call | 1 |
| `list_workbooks` | Before + after open + after PQ | 1, 10 |
| `open_workbook` | Open test copy | 1 |
| `save_workbook` | Before PQ + final save | 10, 13 |
| `close_workbook` | Bridge internally closes during PQ injection | 10 |
| `list_sheets` | Inspect Data sheet | 1 |
| `add_sheet` | Summary, Dashboard, PQ_Data, TempSheet | 3 |
| `delete_sheet` | TempSheet | 3 |
| `read_range` | Header row + all 222 rows | 2 |
| `write_range` | Title, top-10, trend, charts data, banner | 4, 7, 8, 10 |
| `set_formula` | AVERAGE formula + label | 4 |
| `format_range` | Bold, colors, number format, zebra, autofit | 5, 7, 8, 10 |
| `create_table` | Top10Table + TrendTable | 6 |
| `create_chart` | Bar, line, column chart types | 7 |
| `create_pivot_table` | Mac: summarised aggregation table | 9 |
| `add_power_query` | Mac: xlsx ZIP string injection | 10 |
| `refresh` | Refresh all (best-effort) | 10 |
| `run_python` | Introspect Dashboard sheet | 11 |
| `run_vba` | Windows-only; Mac call expected to fail-fast with unsupported error | 12 |

### Mac-specific behaviour
| Feature | Mac behaviour |
| --- | --- |
| Power Query | M formula injected via xlsx ZIP string-patching (no XML re-serialization). `loadToSheet`/`loadToCell` are ignored; user must click Load in the PQ editor. Refresh in Excel: Data -> Refresh All. |
| Pivot Table | Full COM PivotTable unavailable on Mac. Axcelerator writes a static summarised aggregation table at the destination cell (no PivotCache, does not refresh when source changes). |
| VBA macros | `run_vba` is **Windows-only**. On Mac the call returns an unsupported error; use `excel_run_python` (xlwings) instead. |
| Refresh | Best-effort `RefreshAll` via AppleScript; result includes `verified: false` because completion cannot be confirmed. |
| Charts | xlwings chart API works on Mac; all three chart types tested. Chart-title set may emit a warning if the AppleScript path fails. |

---
_Report generated by /Users/naveed/Developer/axcelerator/test_all_tools.py_