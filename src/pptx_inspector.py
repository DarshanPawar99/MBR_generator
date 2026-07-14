#!/usr/bin/env python3
"""Inspect slide objects inside a PowerPoint .pptx file.

The utility intentionally uses only the Python standard library so it can run
before presentation-generation dependencies are chosen. It reads the Open XML
parts in a .pptx archive and reports each slide element's name, inferred type,
text, chart/table/image details, geometry, and whether it is likely editable.
"""
from __future__ import annotations

import argparse
import json
import posixpath
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
EMU_PER_INCH = 914400

CHART_TYPES = {
    "areaChart": "Area chart",
    "barChart": "Bar/column chart",
    "bubbleChart": "Bubble chart",
    "doughnutChart": "Doughnut chart",
    "lineChart": "Line chart",
    "ofPieChart": "Pie-of-pie chart",
    "pieChart": "Pie chart",
    "radarChart": "Radar chart",
    "scatterChart": "Scatter chart",
    "surfaceChart": "Surface chart",
}

@dataclass
class ElementReport:
    slide_number: int
    shape_name: str
    shape_type: str
    current_text: str
    chart_type: str | None
    table_dimensions: str | None
    image_count: int
    position: dict[str, float | None]
    size: dict[str, float | None]
    editable: bool


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _inches(value: str | None) -> float | None:
    if value is None:
        return None
    return round(int(value) / EMU_PER_INCH, 3)


def _text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//a:t", NS)).strip()


def _rels(zf: zipfile.ZipFile, slide_path: str) -> dict[str, str]:
    rel_path = posixpath.join(posixpath.dirname(slide_path), "_rels", posixpath.basename(slide_path) + ".rels")
    if rel_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rel_path))
    out: dict[str, str] = {}
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if rid:
            out[rid] = posixpath.normpath(posixpath.join(posixpath.dirname(slide_path), target))
    return out


def _chart_type(zf: zipfile.ZipFile, chart_path: str) -> str:
    try:
        root = ET.fromstring(zf.read(chart_path))
    except KeyError:
        return "Unknown chart"
    for elem in root.iter():
        name = _local(elem.tag)
        if name in CHART_TYPES:
            return CHART_TYPES[name]
    return "Unknown chart"


def _geometry(node: ET.Element) -> tuple[dict[str, float | None], dict[str, float | None]]:
    xfrm = node.find(".//a:xfrm", NS)
    if xfrm is None:
        xfrm = node.find(".//p:xfrm", NS)
    off = xfrm.find("a:off", NS) if xfrm is not None else None
    ext = xfrm.find("a:ext", NS) if xfrm is not None else None
    return (
        {"left_in": _inches(off.attrib.get("x") if off is not None else None), "top_in": _inches(off.attrib.get("y") if off is not None else None)},
        {"width_in": _inches(ext.attrib.get("cx") if ext is not None else None), "height_in": _inches(ext.attrib.get("cy") if ext is not None else None)},
    )


def _name(node: ET.Element) -> str:
    c_nv_pr = node.find(".//p:cNvPr", NS)
    return c_nv_pr.attrib.get("name", "(unnamed)") if c_nv_pr is not None else "(unnamed)"


def _table_dimensions(node: ET.Element) -> str | None:
    tbl = node.find(".//a:tbl", NS)
    if tbl is None:
        return None
    rows = tbl.findall("a:tr", NS)
    cols = 0
    if rows:
        cols = max(len(row.findall("a:tc", NS)) for row in rows)
    return f"{len(rows)}x{cols}"


def _classify(node: ET.Element, zf: zipfile.ZipFile, rels: dict[str, str]) -> tuple[str, str | None, str | None, int, bool]:
    if node.find(".//a:tbl", NS) is not None:
        return "Table", None, _table_dimensions(node), 0, True
    chart = node.find(".//c:chart", NS)
    if chart is not None:
        rid = chart.attrib.get(f"{{{NS['r']}}}id")
        chart_path = rels.get(rid or "", "")
        return "Chart", _chart_type(zf, chart_path), None, 0, True
    if _local(node.tag) == "pic" or node.find(".//a:blip", NS) is not None:
        return "Image", None, None, 1, False
    if _text(node):
        return "Text", None, None, 0, True
    if _local(node.tag) == "graphicFrame":
        return "Graphic frame", None, None, 0, True
    return "Shape", None, None, 0, True


def _slide_paths(zf: zipfile.ZipFile) -> list[str]:
    paths = [p for p in zf.namelist() if p.startswith("ppt/slides/slide") and p.endswith(".xml")]
    return sorted(paths, key=lambda p: int(Path(p).stem.replace("slide", "")))


def inspect_pptx(path: Path) -> list[ElementReport]:
    reports: list[ElementReport] = []
    with zipfile.ZipFile(path) as zf:
        for slide_number, slide_path in enumerate(_slide_paths(zf), start=1):
            root = ET.fromstring(zf.read(slide_path))
            rels = _rels(zf, slide_path)
            sp_tree = root.find(".//p:spTree", NS)
            if sp_tree is None:
                continue
            for child in list(sp_tree):
                if _local(child.tag) in {"nvGrpSpPr", "grpSpPr"}:
                    continue
                shape_type, chart_type, table_dims, image_count, editable = _classify(child, zf, rels)
                position, size = _geometry(child)
                reports.append(ElementReport(slide_number, _name(child), shape_type, _text(child), chart_type, table_dims, image_count, position, size, editable))
    return reports


def _print_text(reports: Iterable[ElementReport]) -> None:
    current = None
    for item in reports:
        if item.slide_number != current:
            current = item.slide_number
            print(f"Slide {current}")
        detail = item.chart_type or (f"Table {item.table_dimensions}" if item.table_dimensions else "")
        text = f' "{item.current_text}"' if item.current_text else ""
        pos = f" at ({item.position['left_in']}, {item.position['top_in']}) in"
        size = f" size ({item.size['width_in']} x {item.size['height_in']}) in"
        edit = "editable" if item.editable else "not editable"
        print(f"  - {item.shape_name:<28} {item.shape_type:<13} {detail}{text}{pos}{size}; {edit}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect shapes, charts, tables, and images in a .pptx file.")
    parser.add_argument("pptx", type=Path, help="Path to a PowerPoint .pptx file")
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON instead of text")
    args = parser.parse_args(argv)
    if not args.pptx.exists():
        parser.error(f"file not found: {args.pptx}")
    reports = inspect_pptx(args.pptx)
    if args.json:
        print(json.dumps([asdict(r) for r in reports], indent=2))
    else:
        _print_text(reports)
    return 0

if __name__ == "__main__":
    sys.exit(main())
