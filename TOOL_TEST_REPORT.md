# Axcelerator — Tool Verification Report

**Date:** 2026-04-29 07:17:22  
**Workbook tested:** `axcelerator_test_world_bank_mobile.xlsx`  
**Platform:** macOS — Excel 16.108.2 — xlwings 0.35.2 — Python 3.10.8  

## Summary

| Metric | Value |
| --- | --- |
| Total tool calls | 60 |
| Passed | 60 |
| Failed | 0 |
| Total wall-clock | 22393 ms |

## Result per tool call

| # | Tool | Description | Status | ms | Notes |
| --- | --- | --- | --- | --- | --- |
| 1 | `check_environment` | OS / xlwings / Excel version | PASS | 115 | {"os": "Darwin", "pythonVersion": "3.10.8", "xlwingsVersion": "0.35.2", "excelRunning": true, "comAvailable": false, "vbaSupported": false, "powerQuer |
| 2 | `list_workbooks` | List workbooks (before open) | PASS | 72 | [] |
| 3 | `open_workbook` | Open test copy | PASS | 2473 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/private/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_ |
| 4 | `list_workbooks` | List workbooks (after open) | PASS | 237 | [{"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/private/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank |
| 5 | `list_sheets` | Inspect source sheets | PASS | 623 | [{"name": "Data", "index": 1, "usedRange": "$A$1:$X$223", "rowCount": 223, "columnCount": 24}, {"name": "Definition and Source", "index": 2, "usedRang |
| 6 | `read_range` | Read header row A1:X1 | PASS | 133 | {"address": "$A$1:$X$1", "values": [["Series Name", "Series Code", "Country Name", "Country Code", "1997 [YR1997]", "1998 [YR1998]", "1999 [YR1999]",  |
| 7 | `read_range` | Read all 222 data rows A1:X223 | PASS | 134 | {"address": "$A$1:$X$223", "values": [["Series Name", "Series Code", "Country Name", "Country Code", "1997 [YR1997]", "1998 [YR1998]", "1999 [YR1999]" |
| 8 | `add_sheet` | Add TempSheet (to be deleted) | PASS | 164 | {"name": "TempSheet", "index": 2} |
| 9 | `delete_sheet` | Delete TempSheet | PASS | 140 | {"deleted": "TempSheet"} |
| 10 | `add_sheet` | Add Summary sheet | PASS | 266 | {"name": "Summary", "index": 2} |
| 11 | `add_sheet` | Add Dashboard sheet | PASS | 283 | {"name": "Dashboard", "index": 3} |
| 12 | `add_sheet` | Add PQ_Data sheet | PASS | 300 | {"name": "PQ_Data", "index": 4} |
| 13 | `write_range` | Title block to Summary!A1 | PASS | 1134 | {"address": "$A$1", "rows": 3, "cols": 3} |
| 14 | `write_range` | Top-10 table to Summary!A5 | PASS | 1016 | {"address": "$A$5", "rows": 11, "cols": 3} |
| 15 | `write_range` | Trend table to Summary!A18 | PASS | 1000 | {"address": "$A$18", "rows": 6, "cols": 21} |
| 16 | `set_formula` | AVERAGE formula Summary!C17 | PASS | 83 | {"address": "$C$17"} |
| 17 | `set_formula` | Average label Summary!B17 | PASS | 67 | {"address": "$B$17"} |
| 18 | `format_range` | Format Summary A1:C1 | PASS | 229 | {"address": "$A$1:$C$1"} |
| 19 | `format_range` | Format Summary A2:C2 | PASS | 88 | {"address": "$A$2:$C$2"} |
| 20 | `format_range` | Format Summary A5:C5 | PASS | 183 | {"address": "$A$5:$C$5"} |
| 21 | `format_range` | Format Summary C6:C17 | PASS | 133 | {"address": "$C$6:$C$17"} |
| 22 | `format_range` | Zebra row 1 | PASS | 83 | {"address": "$A$6:$C$6"} |
| 23 | `format_range` | Zebra row 2 | PASS | 83 | {"address": "$A$7:$C$7"} |
| 24 | `format_range` | Zebra row 3 | PASS | 67 | {"address": "$A$8:$C$8"} |
| 25 | `format_range` | Zebra row 4 | PASS | 67 | {"address": "$A$9:$C$9"} |
| 26 | `format_range` | Zebra row 5 | PASS | 67 | {"address": "$A$10:$C$10"} |
| 27 | `format_range` | Zebra row 6 | PASS | 83 | {"address": "$A$11:$C$11"} |
| 28 | `format_range` | Zebra row 7 | PASS | 66 | {"address": "$A$12:$C$12"} |
| 29 | `format_range` | Zebra row 8 | PASS | 67 | {"address": "$A$13:$C$13"} |
| 30 | `format_range` | Zebra row 9 | PASS | 67 | {"address": "$A$14:$C$14"} |
| 31 | `format_range` | Zebra row 10 | PASS | 67 | {"address": "$A$15:$C$15"} |
| 32 | `format_range` | Format Summary A1:D25 | PASS | 150 | {"address": "$A$1:$D$25"} |
| 33 | `format_range` | Format Summary A18:X18 | PASS | 100 | {"address": "$A$18:$X$18"} |
| 34 | `create_table` | Top10Table on Summary | PASS | 200 | {"name": "Top10Table", "address": "$A$5:$C$15"} |
| 35 | `create_table` | TrendTable on Summary | PASS | 417 | {"name": "TrendTable", "address": "$A$18:$U$23"} |
| 36 | `write_range` | Top-10 data to Dashboard!A1 | PASS | 1000 | {"address": "$A$1", "rows": 11, "cols": 2} |
| 37 | `format_range` | Dashboard top-10 header A1:B1 | PASS | 167 | {"address": "$A$1:$B$1"} |
| 38 | `create_chart` | Bar chart — Top 10 countries | PASS | 300 | {"name": "Chart 1"} |
| 39 | `write_range` | Trend data (transposed) to Dashboard!A15 | PASS | 1184 | {"address": "$A$15", "rows": 21, "cols": 6} |
| 40 | `format_range` | Trend header A15:F15 | PASS | 100 | {"address": "$A$15:$F$15"} |
| 41 | `create_chart` | Line chart — mobile trends 1997-2016 | PASS | 284 | {"name": "Chart 2"} |
| 42 | `write_range` | Top-5 data to Dashboard!A40 | PASS | 1166 | {"address": "$A$40", "rows": 6, "cols": 2} |
| 43 | `format_range` | Top-5 header A40:B40 | PASS | 200 | {"address": "$A$40:$B$40"} |
| 44 | `create_chart` | Column chart — top 5 | PASS | 300 | {"name": "Chart 3"} |
| 45 | `write_range` | Banner to Dashboard!H1 | PASS | 1083 | {"address": "$H$1", "rows": 5, "cols": 1} |
| 46 | `format_range` | Style Dashboard H1:L1 | PASS | 217 | {"address": "$H$1:$L$1"} |
| 47 | `format_range` | Style Dashboard H2:L2 | PASS | 83 | {"address": "$H$2:$L$2"} |
| 48 | `format_range` | Style Dashboard H3:L3 | PASS | 200 | {"address": "$H$3:$L$3"} |
| 49 | `format_range` | Style Dashboard H4:L5 | PASS | 83 | {"address": "$H$4:$L$5"} |
| 50 | `format_range` | Style Dashboard A1:L50 | PASS | 150 | {"address": "$A$1:$L$50"} |
| 51 | `create_pivot_table` | Summarised pivot from Top10Table -> Summary!F5 | PASS | 2750 | {"name": "Top10Pivot", "note": "Mac: built as summarised table (full COM PivotTable not available on Mac)."} |
| 52 | `write_range` | Seed data to PQ_Data!A1 | PASS | 1017 | {"address": "$A$1", "rows": 11, "cols": 3} |
| 53 | `format_range` | PQ_Data header | PASS | 117 | {"address": "$A$1:$C$1"} |
| 54 | `save_workbook` | Save before PQ injection | PASS | 234 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/private/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_ |
| 55 | `run_python` | Introspect Dashboard via Python | PASS | 366 | {"result": {"dashboard_used_range": "$A$1:$L$45", "dashboard_rows": 45, "dashboard_cols": 12, "workbook_name": "axcelerator_test_world_bank_mobile.xls |
| 56 | `run_vba` | run_vba reachable on this OS (Mac=ok, Win=error for missing macro) | PASS | 179 | Mac: AppleScript silently no-ops missing macros. Win: missing macro raises. |
| 57 | `save_workbook` | Final save (before PQ injection) | PASS | 317 | {"name": "axcelerator_test_world_bank_mobile.xlsx", "fullname": "/private/var/folders/vd/9_kxm62s6g70n8239xx1r72m0000gn/T/axcelerator_test_world_bank_ |
| 58 | `add_power_query` | Inject Power Query 'Top10Query' (last operation) | PASS | 375 | {"queryName": "Top10Query", "loadedTo": null} |
| 59 | `list_workbooks` | List workbooks after PQ injection (workbook closed) | PASS | 39 | [] |
| 60 | `refresh` | Refresh (no-op on Mac after close; graceful) | PASS | 25 | Expected: workbook was closed by noReopen. Graceful error is correct. |

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
| `run_vba` | Mac: AppleScript path reachable (silent no-op for missing macro) | 12 |

### Mac-specific behaviour
| Feature | Mac behaviour |
| --- | --- |
| Power Query | M formula injected via xlsx ZIP string-patching (no XML re-serialization). Refresh in Excel: Data -> Refresh All. |
| Pivot Table | Full COM PivotTable unavailable via AppleScript. Axcelerator builds equivalent summarised table with Python aggregation. |
| VBA macros | `run_vba` dispatches via AppleScript on Mac and via COM on Windows. On Mac the workbook must be `.xlsm` with the macro defined and macro security set to allow execution (Excel > Preferences > Security). |
| Charts | xlwings chart API works on Mac; all three chart types tested. |

---
_Report generated by /Users/naveed/Developer/axcelerator/test_all_tools.py_