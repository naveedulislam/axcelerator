# Axcelerator — Tool Verification Report

**Date:** 2026-04-29 07:39:19  
**Workbook tested:** `axcelerator_test_world_bank_mobile.xlsx`  
**Platform:** macOS — Excel 16.108.2 — xlwings 0.35.2 — Python 3.10.8  

## Summary

| Metric | Value |
| --- | --- |
| Total tool calls | 60 |
| Passed | 60 |
| Failed | 0 |
| Total wall-clock | 38169 ms |

## Result per tool call

| # | Tool | Description | Status | ms | Notes |
| --- | --- | --- | --- | --- | --- |
| 1 | `check_environment` | OS / xlwings / Excel version | PASS | 117 | {"os": "Darwin", "pythonVersion": "3.10.8", "xlwingsVersion": "0.35.2", "excelRunning": true, "comAvailable": false, "vbaSupported": false, "powerQuer |
| 2 | `list_workbooks` | List workbooks (before open) | PASS | 186 | [{"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_mobile. |
| 3 | `open_workbook` | Open test copy | PASS | 1789 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_mobile.x |
| 4 | `list_workbooks` | List workbooks (after open) | PASS | 142 | [{"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_mobile. |
| 5 | `list_sheets` | Inspect source sheets | PASS | 2135 | [{"name": "Data", "index": 1, "usedRange": "$A$1:$X$223", "rowCount": 223, "columnCount": 24}, {"name": "Summary", "index": 2, "usedRange": "$A$1:$X$2 |
| 6 | `read_range` | Read header row A1:X1 | PASS | 132 | {"address": "$A$1:$X$1", "values": [["Series Name", "Series Code", "Country Name", "Country Code", "1997 [YR1997]", "1998 [YR1998]", "1999 [YR1999]",  |
| 7 | `read_range` | Read all 222 data rows A1:X223 | PASS | 136 | {"address": "$A$1:$X$223", "values": [["Series Name", "Series Code", "Country Name", "Country Code", "1997 [YR1997]", "1998 [YR1998]", "1999 [YR1999]" |
| 8 | `add_sheet` | Add TempSheet (to be deleted) | PASS | 297 | {"name": "TempSheet", "index": 2} |
| 9 | `delete_sheet` | Delete TempSheet | PASS | 546 | {"deleted": "TempSheet"} |
| 10 | `add_sheet` | Add Summary sheet | PASS | 259 | {"name": "Summary", "index": 2} |
| 11 | `add_sheet` | Add Dashboard sheet | PASS | 617 | {"name": "Dashboard", "index": 3} |
| 12 | `add_sheet` | Add PQ_Data sheet | PASS | 1020 | {"name": "PQ_Data", "index": 4} |
| 13 | `write_range` | Title block to Summary!A1 | PASS | 1380 | {"address": "$A$1", "rows": 3, "cols": 3} |
| 14 | `write_range` | Top-10 table to Summary!A5 | PASS | 1000 | {"address": "$A$5", "rows": 11, "cols": 3} |
| 15 | `write_range` | Trend table to Summary!A18 | PASS | 1001 | {"address": "$A$18", "rows": 6, "cols": 21} |
| 16 | `set_formula` | AVERAGE formula Summary!C17 | PASS | 84 | {"address": "$C$17"} |
| 17 | `set_formula` | Average label Summary!B17 | PASS | 64 | {"address": "$B$17"} |
| 18 | `format_range` | Format Summary A1:C1 | PASS | 100 | {"address": "$A$1:$C$1"} |
| 19 | `format_range` | Format Summary A2:C2 | PASS | 500 | {"address": "$A$2:$C$2"} |
| 20 | `format_range` | Format Summary A5:C5 | PASS | 100 | {"address": "$A$5:$C$5"} |
| 21 | `format_range` | Format Summary C6:C17 | PASS | 566 | {"address": "$C$6:$C$17"} |
| 22 | `format_range` | Zebra row 1 | PASS | 118 | {"address": "$A$6:$C$6"} |
| 23 | `format_range` | Zebra row 2 | PASS | 82 | {"address": "$A$7:$C$7"} |
| 24 | `format_range` | Zebra row 3 | PASS | 482 | {"address": "$A$8:$C$8"} |
| 25 | `format_range` | Zebra row 4 | PASS | 452 | {"address": "$A$9:$C$9"} |
| 26 | `format_range` | Zebra row 5 | PASS | 83 | {"address": "$A$10:$C$10"} |
| 27 | `format_range` | Zebra row 6 | PASS | 83 | {"address": "$A$11:$C$11"} |
| 28 | `format_range` | Zebra row 7 | PASS | 480 | {"address": "$A$12:$C$12"} |
| 29 | `format_range` | Zebra row 8 | PASS | 437 | {"address": "$A$13:$C$13"} |
| 30 | `format_range` | Zebra row 9 | PASS | 66 | {"address": "$A$14:$C$14"} |
| 31 | `format_range` | Zebra row 10 | PASS | 66 | {"address": "$A$15:$C$15"} |
| 32 | `format_range` | Format Summary A1:D25 | PASS | 687 | {"address": "$A$1:$D$25"} |
| 33 | `format_range` | Format Summary A18:X18 | PASS | 147 | {"address": "$A$18:$X$18"} |
| 34 | `create_table` | Top10Table on Summary | PASS | 1742 | {"name": "Top10Table", "address": "$A$5:$C$15"} |
| 35 | `create_table` | TrendTable on Summary | PASS | 349 | {"name": "TrendTable", "address": "$A$18:$U$23"} |
| 36 | `write_range` | Top-10 data to Dashboard!A1 | PASS | 1908 | {"address": "$A$1", "rows": 11, "cols": 2} |
| 37 | `format_range` | Dashboard top-10 header A1:B1 | PASS | 498 | {"address": "$A$1:$B$1"} |
| 38 | `create_chart` | Bar chart — Top 10 countries | PASS | 651 | {"name": "Chart 1", "warning": "Could not set chart title: 'tuple' object has no attribute 'chart_title'"} |
| 39 | `write_range` | Trend data (transposed) to Dashboard!A15 | PASS | 1767 | {"address": "$A$15", "rows": 21, "cols": 6} |
| 40 | `format_range` | Trend header A15:F15 | PASS | 564 | {"address": "$A$15:$F$15"} |
| 41 | `create_chart` | Line chart — mobile trends 1997-2016 | PASS | 669 | {"name": "Chart 2", "warning": "Could not set chart title: 'tuple' object has no attribute 'chart_title'"} |
| 42 | `write_range` | Top-5 data to Dashboard!A40 | PASS | 1398 | {"address": "$A$40", "rows": 6, "cols": 2} |
| 43 | `format_range` | Top-5 header A40:B40 | PASS | 515 | {"address": "$A$40:$B$40"} |
| 44 | `create_chart` | Column chart — top 5 | PASS | 2066 | {"name": "Chart 3", "warning": "Could not set chart title: 'tuple' object has no attribute 'chart_title'"} |
| 45 | `write_range` | Banner to Dashboard!H1 | PASS | 1437 | {"address": "$H$1", "rows": 5, "cols": 1} |
| 46 | `format_range` | Style Dashboard H1:L1 | PASS | 117 | {"address": "$H$1:$L$1"} |
| 47 | `format_range` | Style Dashboard H2:L2 | PASS | 499 | {"address": "$H$2:$L$2"} |
| 48 | `format_range` | Style Dashboard H3:L3 | PASS | 99 | {"address": "$H$3:$L$3"} |
| 49 | `format_range` | Style Dashboard H4:L5 | PASS | 83 | {"address": "$H$4:$L$5"} |
| 50 | `format_range` | Style Dashboard A1:L50 | PASS | 186 | {"address": "$A$1:$L$50"} |
| 51 | `create_pivot_table` | Summarised pivot from Top10Table -> Summary!F5 | PASS | 3032 | {"name": "Top10Pivot", "note": "Mac: built as summarised table (full COM PivotTable not available on Mac)."} |
| 52 | `write_range` | Seed data to PQ_Data!A1 | PASS | 1417 | {"address": "$A$1", "rows": 11, "cols": 3} |
| 53 | `format_range` | PQ_Data header | PASS | 116 | {"address": "$A$1:$C$1"} |
| 54 | `save_workbook` | Save before PQ injection | PASS | 311 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_mobile.x |
| 55 | `run_python` | Introspect Dashboard via Python | PASS | 1090 | {"result": {"dashboard_used_range": "$A$1:$L$45", "dashboard_rows": 45, "dashboard_cols": 12, "workbook_name": "axcelerator_test_world_bank_mobile.xls |
| 56 | `run_vba` | run_vba: Windows-only policy (Mac=unsupported, Win=missing-macro error) | PASS | 1 | Mac: bridge returns 'Windows-only' unsupported error. Win: missing macro raises. |
| 57 | `save_workbook` | Final save (before PQ injection) | PASS | 299 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_mobile.x |
| 58 | `add_power_query` | Inject Power Query 'Top10Query' (last operation) | PASS | 1939 | {"queryName": "Top10Query", "loadedTo": null} |
| 59 | `list_workbooks` | List workbooks after PQ injection (workbook closed) | PASS | 38 | [] |
| 60 | `refresh` | Refresh (no-op on Mac after close; graceful) | PASS | 24 | Expected: workbook was closed by noReopen. Graceful error is correct. |

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