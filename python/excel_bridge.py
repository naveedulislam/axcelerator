"""
Axcelerator – xlwings JSON-RPC bridge
=======================================

A long-running stdin/stdout JSON-RPC server. The VS Code extension spawns this
script once and sends one JSON request per line; this script writes one JSON
response per line.

Request:  {"id": <int|str>, "method": "<name>", "params": { ... }}
Response: {"id": <same>, "ok": true,  "result": <any>}
       or {"id": <same>, "ok": false, "error": "<message>", "trace": "<traceback>"}

The bridge requires `xlwings` and a running Microsoft Excel.
On Windows, xlwings uses COM (pywin32), giving access to the full Excel object
model including Power Query Queries API, PivotCaches, VBA macros, etc.
On macOS, xlwings uses AppleScript. Power Query IS supported on Excel for Mac
(2016+ / Microsoft 365), but there is no AppleScript/COM API to add queries
programmatically. We work around this by:
  1. Saving and closing the workbook via xlwings.
  2. Directly patching the xlsx ZIP (xl/connections.xml + xl/queryTables/)
     using Python's built-in zipfile + ElementTree (openpyxl optional).
  3. Reopening the workbook.
PivotTable creation uses xlwings' own API where possible; on Mac the full COM
path is unavailable, so we use openpyxl as a fallback for basic pivot layouts.
VBA macros are genuinely Windows-only (Excel for Mac removed VBA support).
"""

from __future__ import annotations

import json
import os
import platform
import shutil
import sys
import tempfile
import traceback
import uuid
import zipfile
from typing import Any, Callable, Dict, List, Optional
from xml.etree import ElementTree as ET

try:
    import xlwings as xw
except Exception as exc:  # pragma: no cover - reported back to extension
    sys.stdout.write(json.dumps({
        "id": None, "ok": False,
        "error": f"xlwings is not installed in this Python interpreter: {exc}",
        "trace": ""
    }) + "\n")
    sys.stdout.flush()
    sys.exit(1)


IS_WINDOWS = platform.system() == "Windows"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _norm_path(path: str) -> str:
    """Normalize a filesystem path: expand ~ and $VARS, then make absolute."""
    if not path:
        return path
    return os.path.abspath(os.path.expandvars(os.path.expanduser(path)))


def _get_app(visible: bool = True) -> "xw.App":
    """Return an active Excel App, launching one if needed."""
    if xw.apps.count == 0:
        return xw.App(visible=visible, add_book=False)
    app = xw.apps.active
    if visible:
        try:
            app.visible = True
        except Exception:
            pass
    return app


def _find_workbook(identifier: str) -> "xw.Book":
    """Find an open workbook by name, basename, or full path."""
    if not identifier:
        raise ValueError("workbook identifier is required")
    norm = os.path.normcase(_norm_path(identifier)) if os.sep in identifier or "/" in identifier else None
    for app in xw.apps:
        for book in app.books:
            if book.name == identifier:
                return book
            if norm and os.path.normcase(book.fullname) == norm:
                return book
            if os.path.basename(book.name) == os.path.basename(identifier):
                return book
    raise ValueError(f"Workbook not found among open workbooks: {identifier!r}")


def _get_sheet(book: "xw.Book", sheet: str) -> "xw.Sheet":
    try:
        return book.sheets[sheet]
    except Exception as exc:
        raise ValueError(f"Sheet {sheet!r} not found in workbook {book.name!r}") from exc


def _hex_to_rgb(hex_color: str) -> tuple:
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid hex color: {hex_color!r}")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------------------------------------------------------------------------
# Method handlers
# ---------------------------------------------------------------------------

def m_check_environment(_: dict) -> dict:
    info: Dict[str, Any] = {
        "os": platform.system(),
        "pythonVersion": platform.python_version(),
        "xlwingsVersion": xw.__version__,
        "excelRunning": xw.apps.count > 0,
        # COM is Windows-only; unlocks Queries API, PivotCaches, VBA macros.
        "comAvailable": IS_WINDOWS,
        # VBA is Windows-only in this extension. Mac VBA was previously a
        # best-effort path but proved unreliable across Excel/AppleScript
        # versions and is now refused at the TS gate as well.
        "vbaSupported": IS_WINDOWS,
        # Power Query works on both. Mac path patches the .xlsx XML; the
        # query is visible in the PQ editor but is NOT auto-loaded to a sheet.
        "powerQuerySupported": True,
        "powerQueryMode": "com" if IS_WINDOWS else "xml-patch",
        # PivotTable is supported on both, but only Windows builds a real
        # PivotCache; Mac writes a static summarised table at the destination.
        "pivotTableSupported": True,
        "pivotTableMode": "native" if IS_WINDOWS else "summary-fallback",
        "pivotTableFullCom": IS_WINDOWS,
    }
    if xw.apps.count > 0:
        try:
            info["excelVersion"] = xw.apps.active.version
        except Exception:
            pass
    return info


def m_list_workbooks(_: dict) -> list:
    out = []
    for app in xw.apps:
        for book in app.books:
            try:
                active_sheet = book.sheets.active.name
            except Exception:
                active_sheet = None
            out.append({
                "name": book.name,
                "fullname": book.fullname,
                "saved": getattr(book, "saved", None),
                "activeSheet": active_sheet,
                "appPid": getattr(app, "pid", None),
            })
    return out


def m_open_workbook(p: dict) -> dict:
    path = p.get("path")
    create = bool(p.get("create", False))
    visible = bool(p.get("visible", True))
    app = _get_app(visible=visible)

    if not path:
        book = app.books.add()
        return {"name": book.name, "fullname": book.fullname, "created": True}

    abs_path = _norm_path(path)
    if os.path.exists(abs_path):
        book = app.books.open(abs_path)
        return {"name": book.name, "fullname": book.fullname, "created": False}

    if not create:
        raise FileNotFoundError(f"File does not exist and create=false: {abs_path}")

    book = app.books.add()
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
    book.save(abs_path)
    return {"name": book.name, "fullname": book.fullname, "created": True}


def m_save_workbook(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    path = p.get("path")
    last_exc = None
    for attempt in range(3):
        try:
            if path:
                book.save(_norm_path(path))
            else:
                book.save()
            return {"name": book.name, "fullname": book.fullname}
        except Exception as exc:
            last_exc = exc
            # OSERROR -50 "Parameter error" on Mac = Excel showing a dialog
            # (e.g. "Enable Queries & Connections"). Wait and retry.
            if "OSERROR: -50" in str(exc) or "Parameter error" in str(exc):
                import time as _time
                if attempt < 2:
                    _time.sleep(2)
                    continue
            raise
    raise last_exc


def m_close_workbook(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    if p.get("save"):
        book.save()
    name = book.name
    book.close()
    return {"closed": name}


def m_list_sheets(p: dict) -> list:
    book = _find_workbook(p["workbook"])
    out = []
    for sh in book.sheets:
        used = sh.used_range
        out.append({
            "name": sh.name,
            "index": sh.index,
            "usedRange": used.address if used else None,
            "rowCount": used.rows.count if used else 0,
            "columnCount": used.columns.count if used else 0,
        })
    return out


def m_add_sheet(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    after = p.get("after")
    after_sheet = book.sheets[after] if after else None
    sh = book.sheets.add(name=p["name"], after=after_sheet)
    return {"name": sh.name, "index": sh.index}


def m_delete_sheet(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    sh = _get_sheet(book, p["sheet"])
    name = sh.name
    sh.delete()
    return {"deleted": name}


def m_read_range(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    sh = _get_sheet(book, p["sheet"])
    rng = sh.range(p["range"])
    expand = p.get("expand", "none")
    if expand and expand != "none":
        rng = rng.expand(expand)
    if p.get("formulas"):
        data = rng.formula
    else:
        data = rng.value
    # Normalise to a 2D list for predictability.
    if data is None:
        values = [[None]]
    elif not isinstance(data, list):
        values = [[data]]
    elif data and not isinstance(data[0], list):
        values = [data]
    else:
        values = data
    return {"address": rng.address, "values": values}


def m_write_range(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    sh = _get_sheet(book, p["sheet"])
    rng = sh.range(p["range"])
    rng.value = p["values"]
    return {"address": rng.address, "rows": len(p["values"]),
            "cols": len(p["values"][0]) if p["values"] and isinstance(p["values"][0], list) else 1}


def m_set_formula(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    sh = _get_sheet(book, p["sheet"])
    rng = sh.range(p["range"])
    is_array = bool(p.get("array", False))
    if "formulas" in p and p["formulas"] is not None:
        if is_array:
            rng.formula_array = p["formulas"]
        else:
            rng.formula = p["formulas"]
    elif "formula" in p and p["formula"] is not None:
        if is_array:
            rng.formula_array = p["formula"]
        else:
            rng.formula = p["formula"]
    else:
        raise ValueError("Provide either `formula` or `formulas`.")
    return {"address": rng.address}


def m_format_range(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    sh = _get_sheet(book, p["sheet"])
    rng = sh.range(p["range"])
    if "numberFormat" in p and p["numberFormat"] is not None:
        rng.number_format = p["numberFormat"]
    if "fontColor" in p and p["fontColor"]:
        rng.font.color = _hex_to_rgb(p["fontColor"])
    if "backgroundColor" in p and p["backgroundColor"]:
        rng.color = _hex_to_rgb(p["backgroundColor"])
    if "bold" in p and p["bold"] is not None:
        rng.font.bold = bool(p["bold"])
    if "italic" in p and p["italic"] is not None:
        rng.font.italic = bool(p["italic"])
    if p.get("autofit"):
        rng.autofit()
    if "columnWidth" in p and p["columnWidth"] is not None:
        rng.column_width = float(p["columnWidth"])
    return {"address": rng.address}


def m_create_table(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    sh = _get_sheet(book, p["sheet"])
    rng = sh.range(p["range"])
    name = p["name"]
    style = p.get("tableStyle", "TableStyleMedium2")
    has_headers = bool(p.get("hasHeaders", True))

    if IS_WINDOWS:
        # Use COM directly for full ListObject support.
        list_objects = sh.api.ListObjects
        xl_src_range = 1  # xlSrcRange
        xl_yes = 1
        xl_no = 2
        lo = list_objects.Add(xl_src_range, rng.api, None,
                              xl_yes if has_headers else xl_no)
        lo.Name = name
        try:
            lo.TableStyle = style
        except Exception:
            pass
        return {"name": name, "address": rng.address}
    else:
        # xlwings has a tables collection that works on Mac too.
        try:
            tbl = sh.tables.add(source=rng, name=name, table_style_name=style,
                                has_headers=has_headers)
            return {"name": tbl.name, "address": rng.address}
        except Exception as exc:
            raise RuntimeError(f"Creating tables on this platform failed: {exc}")


_CHART_TYPE_MAP_WIN = {
    "column_clustered": 51,   # xlColumnClustered
    "bar_clustered": 57,      # xlBarClustered
    "line": 4,                # xlLine
    "line_markers": 65,       # xlLineMarkers
    "pie": 5,                 # xlPie
    "area": 1,                # xlArea
    "scatter": -4169,         # xlXYScatter
    "doughnut": -4120,        # xlDoughnut
}


def m_create_chart(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    sh = _get_sheet(book, p["sheet"])
    src = sh.range(p["sourceRange"])
    chart = sh.charts.add(width=p.get("width", 480), height=p.get("height", 300))
    chart.set_source_data(src)
    ctype = p.get("chartType", "column_clustered")
    try:
        # xlwings uses string enum names that match COM constants on Win.
        chart.chart_type = ctype
    except Exception:
        if IS_WINDOWS and ctype in _CHART_TYPE_MAP_WIN:
            chart.api[1].ChartType = _CHART_TYPE_MAP_WIN[ctype]
    if p.get("title"):
        try:
            if IS_WINDOWS:
                chart.api[1].HasTitle = True
                chart.api[1].ChartTitle.Text = p["title"]
            else:
                chart.api.chart_title.set(p["title"])
        except Exception:
            pass
    if p.get("anchorCell"):
        anchor = sh.range(p["anchorCell"])
        chart.left = anchor.left
        chart.top = anchor.top
    return {"name": chart.name}


def m_create_pivot_table(p: dict) -> dict:
    if IS_WINDOWS:
        book = _find_workbook(p["workbook"])
        src_table_name = p["sourceTable"]
        dest_sheet = _get_sheet(book, p["destinationSheet"])
        dest_cell = dest_sheet.range(p.get("destinationCell", "A1"))

        # Locate the source ListObject.
        src_lo = None
        for sh in book.sheets:
            for lo in sh.api.ListObjects:
                if lo.Name == src_table_name:
                    src_lo = lo
                    break
            if src_lo is not None:
                break
        if src_lo is None:
            raise ValueError(f"Source table not found: {src_table_name!r}")

        xl_database = 1  # xlDatabase
        xl_pivot_table_version = 6  # xlPivotTableVersion15+
        cache = book.api.PivotCaches().Create(SourceType=xl_database, SourceData=src_lo.Range,
                                              Version=xl_pivot_table_version)
        pt = cache.CreatePivotTable(TableDestination=dest_cell.api,
                                    TableName=p["name"],
                                    DefaultVersion=xl_pivot_table_version)

        XL_ROW = 1
        XL_COLUMN = 2
        XL_PAGE = 3
        FUNC_MAP = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139}

        for fld in p.get("rows") or []:
            pt.PivotFields(fld).Orientation = XL_ROW
        for fld in p.get("columns") or []:
            pt.PivotFields(fld).Orientation = XL_COLUMN
        for fld in p.get("filters") or []:
            pt.PivotFields(fld).Orientation = XL_PAGE
        for v in p.get("values") or []:
            fname = v["field"]
            func = FUNC_MAP.get(v.get("function", "sum"), -4157)
            pt.AddDataField(pt.PivotFields(fname), f"{v.get('function', 'sum').title()} of {fname}", func)

        return {"name": pt.Name, "destination": dest_cell.address}
    else:
        # On Mac, build a summary pivot-style table using openpyxl as a fallback.
        # Full interactive PivotTable requires the workbook to be saved first.
        book = _find_workbook(p["workbook"])
        src_table_name = p["sourceTable"]
        dest_sht = _get_sheet(book, p["destinationSheet"])

        # Read source table data via xlwings.
        src_sheet = None
        tbl_range = None
        for sh in book.sheets:
            try:
                for tbl in sh.tables:
                    if tbl.name == src_table_name:
                        src_sheet = sh
                        # Use the full table range (header + data) — data_body_range
                        # can return None on Mac when xlwings can't read it via AS.
                        tbl_range = tbl.range
                        break
            except Exception:
                pass
            if src_sheet:
                break
        if src_sheet is None:
            raise ValueError(
                f"Source table {src_table_name!r} not found. On Mac, full COM PivotTable "
                "is not available; this fallback builds a summarised table using Python."
            )

        raw = tbl_range.options(ndim=2).value or []
        if not raw:
            raise ValueError(f"Source table {src_table_name!r} appears empty.")
        headers = [str(h) if h is not None else "" for h in raw[0]]
        rows_data = raw[1:]  # everything after the header row
        # Build {header: index} mapping.
        col_idx = {h: i for i, h in enumerate(headers)}

        row_fields = p.get("rows") or []
        val_specs = p.get("values") or []
        FUNC_MAP_PY = {"sum": sum, "count": len, "average": lambda x: sum(x)/len(x) if x else 0,
                       "max": max, "min": min}

        # Aggregate.
        from collections import defaultdict
        agg: Dict[tuple, Dict[str, list]] = defaultdict(lambda: defaultdict(list))
        for row in rows_data:
            key = tuple(row[col_idx[f]] for f in row_fields if f in col_idx)
            for vs in val_specs:
                fname = vs["field"]
                if fname in col_idx:
                    v = row[col_idx[fname]]
                    if v is not None:
                        agg[key][fname].append(v)

        # Write results to destination sheet.
        anchor_col_letter = p.get("destinationCell", "A1")
        out_headers = row_fields + [f"{vs.get('function','sum').title()} of {vs['field']}" for vs in val_specs]
        dest_sht.range(anchor_col_letter).value = [out_headers]
        result_rows = []
        for key, vals in sorted(agg.items()):
            result_row = list(key)
            for vs in val_specs:
                fn = FUNC_MAP_PY.get(vs.get("function", "sum"), sum)
                result_row.append(fn(vals.get(vs["field"], [0])))
            result_rows.append(result_row)
        if result_rows:
            cell = dest_sht.range(anchor_col_letter).offset(1, 0)
            cell.value = result_rows

        return {"name": p["name"],
                "note": "Mac: built as summarised table (full COM PivotTable not available on Mac)."}


# ---------------------------------------------------------------------------
# Power Query helpers (Mac xlsx-patching path)
# ---------------------------------------------------------------------------

def _xml_escape(s: str) -> str:
    """Minimal XML attribute/text escaping."""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _build_datamashup(queries: list[tuple[str, str]]) -> bytes:
    """
    Build a DataMashup binary package (MS-XLDM format) for Excel Power Query
    on Mac.  Layout (matching what Excel for Mac writes):

        [4 bytes] version         = 0x00000000
        [4 bytes] packageSize
        [N bytes] inner OPC ZIP   (Config/Package.xml, Formulas/Section1.m,
                                    [Content_Types].xml — order matters)
        [4 bytes] permissionsSize
        [P bytes] permissions XML (UTF-8 with BOM)
        [4 bytes] metadataBlockSize
        [Q bytes] metadata block:
                  [4 bytes] 0
                  [4 bytes] xmlSize
                  [X bytes] metadata XML (UTF-8 with BOM)
                  [4 bytes] eocdSize
                  [E bytes] inner-ZIP EOCD-shaped trailer
    """
    import io as _io
    import struct as _struct

    # ----- Section1.m (matches Excel for Mac's exact formatting) -----
    section_parts = ["section Section1;", ""]
    for qname, formula in queries:
        clean = formula.rstrip().rstrip(";").strip()
        section_parts.append(f"shared {qname} = {clean};")
    section_m = "\r\n".join(section_parts) + "\r\n"

    content_types = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="text/xml" />'
        '<Default Extension="m" ContentType="application/x-ms-m" />'
        '</Types>'
    )
    package_xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<Package xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
        ' xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        '<Version>2.153.226.0</Version>'
        '<MinVersion>2.21.0.0</MinVersion>'
        '<Culture>en-US</Culture>'
        '</Package>'
    )

    # Inner OPC ZIP — order observed in Excel-Mac output:
    #   Config/Package.xml, Formulas/Section1.m, [Content_Types].xml
    buf = _io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Config/Package.xml",  package_xml.encode("utf-8"))
        z.writestr("Formulas/Section1.m", section_m.encode("utf-8"))
        z.writestr("[Content_Types].xml", content_types.encode("utf-8"))
    inner_zip = buf.getvalue()

    # ----- Permissions block (BOM + UTF-8 XML) -----
    perms_xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
        ' xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        '<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>'
        '<FirewallEnabled>true</FirewallEnabled>'
        '</PermissionList>'
    )
    perms_bytes = b'\xef\xbb\xbf' + perms_xml.encode("utf-8")

    # ----- Metadata XML (BOM + UTF-8) -----
    item_blocks = [
        '<Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation>'
        '<StableEntries><Entry Type="IsTypeDetectionEnabled" Value="sTrue" /></StableEntries></Item>'
    ]
    for qname, _formula in queries:
        qn_esc = _xml_escape(qname)
        item_blocks.append(
            f'<Item><ItemLocation><ItemType>Formula</ItemType>'
            f'<ItemPath>Section1/{qn_esc}</ItemPath></ItemLocation>'
            f'<StableEntries />'
            f'</Item>'
        )
        item_blocks.append(
            f'<Item><ItemLocation><ItemType>Formula</ItemType>'
            f'<ItemPath>Section1/{qn_esc}/Source</ItemPath></ItemLocation>'
            f'<StableEntries />'
            f'</Item>'
        )

    metadata_xml = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<LocalPackageMetadataFile xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
        ' xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        '<Items>' + "".join(item_blocks) + '</Items>'
        '</LocalPackageMetadataFile>'
    )
    metadata_xml_bytes = b'\xef\xbb\xbf' + metadata_xml.encode("utf-8")

    # Inner-ZIP EOCD trailer (22 bytes, observed in Excel-Mac output)
    eocd = (
        b'PK\x05\x06'        # EOCD signature
        + b'\x00' * 18       # zeroed fields (number of disks, entries, sizes, comment len)
    )

    metadata_block = (
        _struct.pack("<I", 0)                       # padding/version
        + _struct.pack("<I", len(metadata_xml_bytes))
        + metadata_xml_bytes
        + _struct.pack("<I", len(eocd))
        + eocd
    )

    # Outer trailer (100 bytes; Excel-Mac uses a content-derived signature
    # here, but accepts a zero-filled trailer when the workbook was not
    # produced by signed Office tooling).
    outer_trailer = b'\x00' * 100

    return (
        _struct.pack("<I", 0)                       # version
        + _struct.pack("<I", len(inner_zip))
        + inner_zip
        + _struct.pack("<I", len(perms_bytes))
        + perms_bytes
        + _struct.pack("<I", len(metadata_block))
        + metadata_block
        + _struct.pack("<I", len(outer_trailer))
        + outer_trailer
    )


def _wrap_datamashup_xml(blob: bytes) -> bytes:
    """
    Wrap a DataMashup binary blob in the UTF-16 XML envelope that Excel for
    Mac stores in customXml/item1.xml:

        <?xml version="1.0" encoding="utf-16"?>
        <DataMashup xmlns="http://schemas.microsoft.com/DataMashup">BASE64</DataMashup>

    Returned bytes start with the UTF-16 LE BOM (FF FE).
    """
    import base64 as _b64
    b64 = _b64.b64encode(blob).decode("ascii")
    text = (
        '<?xml version="1.0" encoding="utf-16"?>'
        '<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
        + b64 +
        '</DataMashup>'
    )
    # UTF-16 LE with BOM
    return b'\xff\xfe' + text.encode("utf-16-le")


def _pq_patch_xlsx(book_path: str, query_name: str, m_formula: str,
                   load_to_sheet: Optional[str], load_to_cell: str,
                   replace: bool) -> None:
    """
    Inject / replace a Power Query (M formula) into an xlsx file using the
    format Excel for Mac actually reads:

      * customXml/item1.xml    – UTF-16 XML wrapping a base64 DataMashup blob
      * customXml/itemProps1.xml – datastoreItem with DataMashup schemaRef
      * customXml/_rels/item1.xml.rels – customXmlProps relationship
      * Standard "customXml" relationship in xl/_rels/workbook.xml.rels
      * No DataMashup content-type override in [Content_Types].xml
      * No x15:queries extension in workbook.xml

    Excel for Mac does NOT use the binary DataMashup blob directly nor the
    x15:queries metadata; the M formulas live entirely inside the base64
    DataMashup payload of customXml/item1.xml.
    """
    import re as _re
    import uuid as _uuid

    tmp = book_path + ".axcelerator_tmp"
    shutil.copy2(book_path, tmp)

    try:
        with zipfile.ZipFile(book_path, "r") as zin:
            names = zin.namelist()
            entries: Dict[str, bytes] = {n: zin.read(n) for n in names}

        # ----------------------------------------------------------------
        # Locate (or choose) the customXml/itemN.xml that holds DataMashup.
        # Prefer an existing one whose itemProps references DataMashup;
        # otherwise allocate the next free customXml/itemN.xml slot.
        # ----------------------------------------------------------------
        dm_part: Optional[str] = None
        # Look for an existing DataMashup item.
        for n, d in entries.items():
            if _re.fullmatch(r"customXml/item(\d+)\.xml", n):
                head = d[:2]
                if head == b"\xff\xfe" and b"DataMashup" in d[:2000]:
                    dm_part = n
                    break

        if dm_part is None:
            used = [
                int(m.group(1))
                for n in entries
                for m in [_re.fullmatch(r"customXml/item(\d+)\.xml", n)]
                if m
            ]
            idx = max(used, default=0) + 1
            dm_part = f"customXml/item{idx}.xml"

        # Derive the matching itemProps and rels paths.
        idx_match = _re.fullmatch(r"customXml/item(\d+)\.xml", dm_part)
        item_idx = int(idx_match.group(1))
        props_part = f"customXml/itemProps{item_idx}.xml"
        item_rels  = f"customXml/_rels/item{item_idx}.xml.rels"

        # ----------------------------------------------------------------
        # 1.  Read existing queries from the current DataMashup blob (if any)
        # ----------------------------------------------------------------
        existing_queries: list[tuple[str, str]] = []
        if dm_part in entries:
            try:
                import base64 as _b64
                import io as _io
                import struct as _struct
                wrapper = entries[dm_part].decode("utf-16-le", errors="replace")
                m = _re.search(r'<DataMashup[^>]*>([^<]+)</DataMashup>', wrapper)
                if m:
                    blob = _b64.b64decode(m.group(1))
                    pkg_size = _struct.unpack_from("<I", blob, 4)[0]
                    inner = blob[8:8 + pkg_size]
                    with zipfile.ZipFile(_io.BytesIO(inner)) as iz:
                        if "Formulas/Section1.m" in iz.namelist():
                            section_text = iz.read("Formulas/Section1.m").decode("utf-8")
                            for em in _re.finditer(
                                r'shared\s+(\S+?)\s*=\s*(.*?);(?=\s*(?:shared|\Z))',
                                section_text,
                                _re.DOTALL,
                            ):
                                existing_queries.append(
                                    (em.group(1).strip(), em.group(2).strip())
                                )
            except Exception:
                existing_queries = []

        # Build the updated list.
        if replace:
            updated_queries = [(qn, f) for qn, f in existing_queries if qn != query_name]
        else:
            if any(qn == query_name for qn, _ in existing_queries):
                raise ValueError(
                    f"Query already exists: {query_name!r}. Pass replace=true to overwrite."
                )
            updated_queries = list(existing_queries)
        updated_queries.append((query_name, m_formula))

        # ----------------------------------------------------------------
        # 2.  customXml/item1.xml  – UTF-16 XML wrapping base64 blob
        # ----------------------------------------------------------------
        entries[dm_part] = _wrap_datamashup_xml(_build_datamashup(updated_queries))

        # ----------------------------------------------------------------
        # 3.  customXml/itemPropsN.xml – datastoreItem with DataMashup schema
        # ----------------------------------------------------------------
        if props_part not in entries:
            item_id = "{" + str(_uuid.uuid4()).upper() + "}"
            entries[props_part] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n'
                f'<ds:datastoreItem ds:itemID="{item_id}"'
                ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
                '<ds:schemaRefs>'
                '<ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup"/>'
                '</ds:schemaRefs>'
                '</ds:datastoreItem>'
            ).encode("utf-8")

        # ----------------------------------------------------------------
        # 4.  customXml/_rels/itemN.xml.rels – points to itemPropsN.xml
        # ----------------------------------------------------------------
        if item_rels not in entries:
            entries[item_rels] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"'
                f' Target="itemProps{item_idx}.xml"/>'
                '</Relationships>'
            ).encode("utf-8")

        # ----------------------------------------------------------------
        # 5.  [Content_Types].xml – customXmlProperties override for itemProps
        #     (NO DataMashup override; item1.xml uses the default xml type)
        # ----------------------------------------------------------------
        ct_text = entries.get("[Content_Types].xml", b"").decode("utf-8")
        # Remove any stale DataMashup override we may have written previously.
        ct_text = _re.sub(
            r'<Override\b[^>]*ContentType="application/vnd\.ms-excel\.datamashupdefinition"[^>]*/>',
            '', ct_text,
        )
        props_override = (
            f'<Override PartName="/{props_part}"'
            ' ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
        )
        if props_part not in ct_text:
            ct_text = ct_text.replace("</Types>", f"{props_override}</Types>", 1)
        entries["[Content_Types].xml"] = ct_text.encode("utf-8")

        # ----------------------------------------------------------------
        # 6.  xl/_rels/workbook.xml.rels – standard customXml relationship
        #     (NOT the DataMashup type).  Also remove any pre-existing
        #     Excel-Mac stub with Target="NULL" or the wrong DataMashup type.
        # ----------------------------------------------------------------
        DM_REL_TYPE_OLD = "http://schemas.microsoft.com/DataMashup"
        CX_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"

        rels_text = entries.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
        # Drop any old DataMashup-typed relationship.
        rels_text = _re.sub(
            rf'<Relationship\b[^>]*Type="{_re.escape(DM_REL_TYPE_OLD)}"[^>]*/>\s*',
            '', rels_text,
        )
        # Only add a customXml relationship if no existing one already targets our part.
        target_rel = f"../{dm_part}"
        if target_rel not in rels_text:
            used_ids = [int(m) for m in _re.findall(r'Id="rId(\d+)"', rels_text)]
            next_id = max(used_ids, default=0) + 1
            new_rel = (
                f'<Relationship Id="rId{next_id}"'
                f' Type="{CX_REL_TYPE}"'
                f' Target="{target_rel}"/>'
            )
            rels_text = rels_text.replace(
                "</Relationships>", f"{new_rel}</Relationships>", 1
            )
        entries["xl/_rels/workbook.xml.rels"] = rels_text.encode("utf-8")

        # ----------------------------------------------------------------
        # 7.  xl/workbook.xml – strip any stale x15:queries we may have
        #     written previously.  Excel for Mac does not use it.
        # ----------------------------------------------------------------
        wb_text = entries.get("xl/workbook.xml", b"").decode("utf-8")
        wb_text = _re.sub(
            r'<ext\b[^>]*uri="\{56347539-B623-4BE6-A4AB-F3DDA97E5BAA\}"[^>]*>.*?</ext>',
            '', wb_text, flags=_re.DOTALL,
        )
        # Clean up any now-empty <extLst></extLst>
        wb_text = _re.sub(r'<extLst>\s*</extLst>', '', wb_text)
        entries["xl/workbook.xml"] = wb_text.encode("utf-8")

        # ----------------------------------------------------------------
        # Rewrite the xlsx ZIP – [Content_Types].xml first per OPC spec.
        # ----------------------------------------------------------------
        with zipfile.ZipFile(book_path, "w", zipfile.ZIP_DEFLATED) as zout:
            if "[Content_Types].xml" in entries:
                zout.writestr("[Content_Types].xml", entries["[Content_Types].xml"])
            for entry_name, data in entries.items():
                if entry_name == "[Content_Types].xml":
                    continue
                zout.writestr(entry_name, data)

        os.remove(tmp)
    except Exception:
        shutil.copy2(tmp, book_path)
        os.remove(tmp)
        raise


def m_add_power_query(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    name = p["queryName"]
    formula = p["mFormula"]
    replace = bool(p.get("replace", True))
    load_to = None

    if IS_WINDOWS:
        queries = book.api.Queries
        existing = None
        for i in range(1, queries.Count + 1):
            q = queries.Item(i)
            if q.Name == name:
                existing = q
                break

        if existing is not None:
            if not replace:
                raise ValueError(f"Query already exists: {name!r}. Pass replace=true to overwrite.")
            existing.Formula = formula
        else:
            queries.Add(Name=name, Formula=formula)

        if p.get("loadToSheet"):
            sheet = _get_sheet(book, p["loadToSheet"])
            cell = sheet.range(p.get("loadToCell", "A1"))
            conn_str = (
                f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;"
                f"Data Source=$Workbook$;Location={name};Extended Properties=\"\""
            )
            cmd_text = f"SELECT * FROM [{name}]"
            wb_conns = book.api.Connections
            conn_name = f"Query - {name}"
            conn = None
            for i in range(1, wb_conns.Count + 1):
                c = wb_conns.Item(i)
                if c.Name == conn_name:
                    conn = c
                    break
            if conn is None:
                conn = wb_conns.Add2(Name=conn_name, Description="",
                                     ConnectionString=conn_str,
                                     CommandText=cmd_text, lCmdtype=2)
            dest = sheet.api.ListObjects.Add(SourceType=0, Source=conn,
                                             LinkSource=True, Destination=cell.api).QueryTable
            dest.CommandType = 2
            dest.CommandText = cmd_text
            dest.Refresh(BackgroundQuery=False)
            load_to = f"{sheet.name}!{cell.address}"
    else:
        # Mac path: save the workbook, close it, patch the xlsx, reopen.
        book_path = book.fullname
        if not book_path or not book_path.endswith((".xlsx", ".xlsm", ".xlam")):
            raise RuntimeError(
                "The workbook must be saved as .xlsx/.xlsm before Power Query can be "
                "injected on Mac. Save it first with excel_save_workbook."
            )
        book.save()
        book.close()
        _pq_patch_xlsx(book_path, name, formula,
                       p.get("loadToSheet"), p.get("loadToCell", "A1"), replace)
        # Reopen so the workbook stays visible (unless caller opts out).
        if not p.get("noReopen", False):
            import time as _time
            app = _get_app(visible=True)
            book = app.books.open(book_path)
            _time.sleep(0.5)
        if p.get("loadToSheet"):
            load_to = f"{p['loadToSheet']}!{p.get('loadToCell', 'A1')} (refresh manually in Excel to load data)"

    return {"queryName": name, "loadedTo": load_to}


def m_refresh(p: dict) -> dict:
    book = _find_workbook(p["workbook"])
    qn = p.get("queryName")
    if IS_WINDOWS:
        if qn:
            for sh in book.sheets:
                for lo in sh.api.ListObjects:
                    try:
                        if lo.QueryTable is not None and lo.Name == qn:
                            lo.QueryTable.Refresh(BackgroundQuery=False)
                            return {"refreshed": qn}
                    except Exception:
                        pass
            raise ValueError(f"Query/table not found: {qn!r}")
        book.api.RefreshAll()
    else:
        # On Mac, RefreshAll via AppleScript.
        try:
            book.api.refresh_all()
        except Exception:
            pass  # May not be available in older xlwings builds; best-effort.
    return {"refreshed": qn or "all"}


def m_run_vba(p: dict) -> dict:
    if not IS_WINDOWS:
        raise RuntimeError(
            "excel_run_vba is Windows-only. Mac VBA via AppleScript is "
            "unreliable across Excel versions and is intentionally not "
            "supported by this extension. Use excel_run_python instead."
        )
    book = _find_workbook(p["workbook"])
    args = p.get("args") or []
    macro = book.macro(p["macro"])
    res = macro(*args)
    return {"result": res}


def m_run_python(p: dict) -> dict:
    code = p["code"]
    book = None
    if p.get("workbook"):
        book = _find_workbook(p["workbook"])
    app = xw.apps.active if xw.apps.count else None
    scope: Dict[str, Any] = {"xw": xw, "app": app, "wb": book, "result": None}
    exec(compile(code, "<excelerator-snippet>", "exec"), scope, scope)
    res = scope.get("result")
    # Best-effort JSON-safe conversion.
    try:
        json.dumps(res)
        return {"result": res}
    except TypeError:
        return {"result": repr(res)}


METHODS: Dict[str, Callable[[dict], Any]] = {
    "check_environment": m_check_environment,
    "list_workbooks": m_list_workbooks,
    "open_workbook": m_open_workbook,
    "save_workbook": m_save_workbook,
    "close_workbook": m_close_workbook,
    "list_sheets": m_list_sheets,
    "add_sheet": m_add_sheet,
    "delete_sheet": m_delete_sheet,
    "read_range": m_read_range,
    "write_range": m_write_range,
    "set_formula": m_set_formula,
    "format_range": m_format_range,
    "create_table": m_create_table,
    "create_chart": m_create_chart,
    "create_pivot_table": m_create_pivot_table,
    "add_power_query": m_add_power_query,
    "refresh": m_refresh,
    "run_vba": m_run_vba,
    "run_python": m_run_python,
}


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def _write(obj: dict) -> None:
    sys.stdout.write(json.dumps(obj, default=str) + "\n")
    sys.stdout.flush()


def main() -> None:
    _write({"id": None, "ok": True, "result": {"ready": True, "methods": sorted(METHODS)}})
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        req_id: Optional[Any] = None
        try:
            req = json.loads(line)
            req_id = req.get("id")
            method = req.get("method")
            params = req.get("params") or {}
            if method not in METHODS:
                raise ValueError(f"Unknown method: {method!r}")
            result = METHODS[method](params)
            _write({"id": req_id, "ok": True, "result": result})
        except Exception as exc:
            _write({
                "id": req_id, "ok": False,
                "error": str(exc),
                "trace": traceback.format_exc(),
            })


if __name__ == "__main__":
    main()
