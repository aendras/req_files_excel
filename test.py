
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from pathlib import PurePosixPath
import re
import json
from html import escape

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

DRAW_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
}


EXCEL_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

DATA_DIR = Path("data")


def safe_read(z, name):
    if name in z.namelist():
        return z.read(name)
    return None


def discover_excel_sheets(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as z:

        if "xl/workbook.xml" not in z.namelist():
            raise ValueError("Missing xl/workbook.xml")

        workbook_root = ET.fromstring(z.read("xl/workbook.xml"))

        sheets = []
        for sheet in workbook_root.findall(".//main:sheet", EXCEL_NS):
            sheets.append({
                "sheet_name": sheet.attrib["name"],
                "sheet_id": sheet.attrib["sheetId"],
                "rId": sheet.attrib[
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                ]
            })

        if "xl/_rels/workbook.xml.rels" not in z.namelist():
            raise ValueError("Missing workbook relationships")

        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))

        # --- READ TARGET + TYPE ---
        rels = {}
        for rel in rels_root.findall(".//"):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            rtype = rel.attrib.get("Type")

            if rid and target and rtype:
                rels[rid] = {
                    "target": target,
                    "type": rtype
                }

        valid_sheets = []

        for s in sheets:
            rel = rels.get(s["rId"])

            if not rel:
                # no relationship → skip
                continue

            if not rel["type"].endswith("/worksheet"):
                # chartsheet / dialog / navigation → skip
                continue

            s["xml_path"] = "xl/" + rel["target"]
            valid_sheets.append(s)

        return valid_sheets


def parse_sheet_metadata(root):
    dim = root.find("x:dimension", NS)
    return {
        "dimension": dim.attrib["ref"] if dim is not None else None
    }


def load_shared_strings(z):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    strings = []
    for si in root.findall("x:si", NS):
        texts = [t.text for t in si.findall(".//x:t", NS) if t.text]
        strings.append("".join(texts))
    return strings



def parse_sheet_data(root, shared_strings):
    rows = []

    for row in root.findall(".//x:row", NS):
        cells = {}
        for cell in row.findall("x:c", NS):
            ref = cell.attrib["r"]
            t = cell.attrib.get("t")
            v = cell.find("x:v", NS)

            if v is None:
                value = None
            elif t == "s":
                value = shared_strings[int(v.text)]
            else:
                value = v.text

            cells[ref] = value

        rows.append({
            "row": int(row.attrib["r"]),
            "cells": cells
        })

    return rows


def parse_merged_cells(root):
    return [
        mc.attrib["ref"]
        for mc in root.findall(".//x:mergeCell", NS)
    ]


def extract_sheet_index(sheet_path):
    m = re.search(r"sheet(\d+)\.xml", sheet_path)
    return int(m.group(1)) if m else None


def parse_drawings_with_anchors(z, sheet_root, sheet_idx):
    drawings = []

    drawing_elem = sheet_root.find("x:drawing", NS)
    if drawing_elem is None:
        return drawings

    rels_path = f"xl/worksheets/_rels/sheet{sheet_idx}.xml.rels"
    if rels_path not in z.namelist():
        return drawings

    rid = drawing_elem.attrib[
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    ]

    rels_root = ET.fromstring(z.read(rels_path))

    target = None
    for r in rels_root:
        if r.attrib.get("Id") == rid:
            target = r.attrib.get("Target")
            break

    if not target:
        return drawings

    drawing_path = str(
        PurePosixPath("xl/worksheets") / target
    ).replace("xl/worksheets/../", "xl/")

    if drawing_path not in z.namelist():
        return drawings

    droot = ET.fromstring(z.read(drawing_path))

    for anchor in droot.findall(".//xdr:twoCellAnchor", DRAW_NS):
        fr = anchor.find("xdr:from", DRAW_NS)
        to = anchor.find("xdr:to", DRAW_NS)

        texts = [
            t.text for t in anchor.findall(".//a:t", DRAW_NS) if t.text
        ]

        drawings.append({
            "from_row": int(fr.find("xdr:row", DRAW_NS).text),
            "to_row": int(to.find("xdr:row", DRAW_NS).text),
            "text": texts
        })

    return drawings

def col_to_idx(col):
    """A -> 0, B -> 1, Z -> 25, AA -> 26"""
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1


def idx_to_col(idx):
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(r + 65) + s
    return s


def split_cell_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


def parse_range(rng):
    start, end = rng.split(":")
    sc, sr = split_cell_ref(start)
    ec, er = split_cell_ref(end)
    return (
        sr, er,
        col_to_idx(sc), col_to_idx(ec)
    )


def build_grid(rows):
    if not rows:
        return []

    max_row = max(r["row"] for r in rows)
    max_col = 0

    for r in rows:
        for ref in r["cells"]:
            col, _ = split_cell_ref(ref)
            max_col = max(max_col, col_to_idx(col))

    grid = [
        ["" for _ in range(max_col + 1)]
        for _ in range(max_row)
    ]

    for r in rows:
        r_idx = r["row"] - 1
        for ref, val in r["cells"].items():
            col, _ = split_cell_ref(ref)
            c_idx = col_to_idx(col)
            grid[r_idx][c_idx] = val or ""

    return grid


def apply_merges(grid, merged_cells):
    spans = {}
    skip = set()

    for rng in merged_cells:
        r1, r2, c1, c2 = parse_range(rng)
        rowspan = r2 - r1 + 1
        colspan = c2 - c1 + 1

        spans[(r1-1, c1)] = {
            "rowspan": rowspan,
            "colspan": colspan
        }

        for r in range(r1-1, r2):
            for c in range(c1, c2+1):
                if (r, c) != (r1-1, c1):
                    skip.add((r, c))

    return spans, skip


def grid_to_html(grid, spans, skip):
    html = ["<table border='1'>"]

    for r, row in enumerate(grid):
        html.append("<tr>")
        for c, val in enumerate(row):
            if (r, c) in skip:
                continue

            attrs = ""
            if (r, c) in spans:
                sp = spans[(r, c)]
                if sp["rowspan"] > 1:
                    attrs += f' rowspan="{sp["rowspan"]}"'
                if sp["colspan"] > 1:
                    attrs += f' colspan="{sp["colspan"]}"'

            html.append(f"<td{attrs}>{escape(str(val))}</td>")
        html.append("</tr>")

    html.append("</table>")
    return "".join(html)


sheet_jsons = []
tables = []

for xlsx in DATA_DIR.glob("*.xlsx"):
    if xlsx.name.startswith("~$"):
        continue

    try:
        sheets = discover_excel_sheets(xlsx)

        with zipfile.ZipFile(xlsx) as z:
            shared_strings = load_shared_strings(z)

            for s in sheets:
                sheet_path = s["xml_path"]
                sheet_name = s["sheet_name"]
                sheet_idx = extract_sheet_index(sheet_path)

                try:
                    if sheet_path not in z.namelist():
                        continue

                    sheet_root = ET.fromstring(z.read(sheet_path))

                    # optional: skip sheets with no real data
                    if not sheet_root.findall(".//x:sheetData", NS):
                        continue

                    sheet_jsons.append({
                        "xlsx_name": xlsx.name,
                        "sheet_name": sheet_name,
                        "sheet_path": sheet_path,
                        "dimension": parse_sheet_metadata(sheet_root),
                        "rows": parse_sheet_data(sheet_root, shared_strings),
                        "merged_cells": parse_merged_cells(sheet_root),
                        "drawings": parse_drawings_with_anchors(
                            z, sheet_root, sheet_idx
                        )
                    })

                    # TODO: append to results list / write to disk

                except Exception as sheet_err:
                    print(
                        f"[SHEET ERROR] {xlsx.name} | {sheet_name}\n"
                        f"  → {sheet_err}"
                    )

    except Exception as file_err:
        print(f"[FILE ERROR] {xlsx.name} → {file_err}")


with open("parsed.json", "w", encoding="utf-8") as f:
    json.dump(sheet_jsons, f, indent=2, ensure_ascii=False)


with open("parsed.json", "r", encoding="utf-8") as f:
    data = json.load(f)


for sheet in data:
    try:
        grid = build_grid(sheet.get("rows", []))
        
        if not grid:
            # No tabular data → skip
            continue
        
        spans, skip = apply_merges(grid, sheet.get("merged_cells", []))
        html = grid_to_html(grid, spans, skip)
        
        tables.append({
            "xlsx": sheet["xlsx_name"],
            "sheet": sheet["sheet_name"],
            "html": html
        })
    except Exception as e:
        print(f" xlsx_name : {sheet["xlsx_name"]}, sheet_name : {sheet["sheet_name"]}, exception : {e}")
tables

with open("tables.json", "w", encoding="utf-8") as f:
    json.dump(tables, f, indent=2, ensure_ascii=False)























