#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้าง Dashboard แผนที่เสาไฟด้วย MapLibre GL JS จากไฟล์ Excel (.xlsx)

จุดเด่น
- อ่านไฟล์ .xlsx ด้วย Python Standard Library (ไม่ต้องติดตั้ง pandas/openpyxl)
- สร้าง HTML ไฟล์เดียว โดยฝังข้อมูลไว้ในไฟล์
- Heatmap จำนวนเหตุการณ์ซ่อม
- มุมมองเสาไฟดับ / วิธีซ่อมสำเร็จ / เสาซ่อมซ้ำ
- ตัวกรอง เขต ชนิดโคม ความเสียหาย อาการ วิธีแก้ไข สถานะ วันที่ และค้นหา
- KPI, อันดับเขต, อันดับวิธีซ่อม และรายละเอียดประวัติเสาที่เลือกจากแผนที่
- Export ข้อมูลที่กรองเป็น CSV

การใช้งาน
    python generate_lighting_dashboard.py --excel test.xlsx --output lighting_dashboard.html

หมายเหตุ
- HTML ต้องเชื่อมต่ออินเทอร์เน็ตเพื่อโหลด MapLibre GL JS และ Basemap OpenFreeMap
- Dashboard ยึดรหัสเสาไฟเป็นหลัก และรวม Ticket ที่อยู่ในช่วงซ่อมเดียวกันเป็นรอบซ่อมเดียว
"""

from __future__ import annotations

import argparse
import json
import math
import re
import statistics
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from html import escape
from pathlib import Path, PurePosixPath
from typing import Any, Iterable
from xml.etree import ElementTree as ET


# -----------------------------
# XLSX reader: Standard Library
# -----------------------------
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def _column_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    if not letters:
        return 0
    value = 0
    for ch in letters.group(0):
        value = value * 26 + (ord(ch) - 64)
    return value - 1


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    path = "xl/sharedStrings.xml"
    if path not in zf.namelist():
        return []
    root = ET.fromstring(zf.read(path))
    strings: list[str] = []
    for si in root.findall(f"{{{NS_MAIN}}}si"):
        texts = [node.text or "" for node in si.iter(f"{{{NS_MAIN}}}t")]
        strings.append("".join(texts))
    return strings


def _workbook_sheet_path(zf: zipfile.ZipFile, sheet_name: str | None) -> tuple[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{NS_REL_PKG}}}Relationship")
    }

    sheets = workbook.find(f"{{{NS_MAIN}}}sheets")
    if sheets is None or not list(sheets):
        raise ValueError("ไม่พบ Worksheet ในไฟล์ Excel")

    selected = None
    available: list[str] = []
    for sheet in sheets:
        name = sheet.attrib.get("name", "")
        available.append(name)
        if sheet_name and name == sheet_name:
            selected = sheet
            break
    if selected is None:
        if sheet_name:
            raise ValueError(f"ไม่พบชีต '{sheet_name}' (มีชีต: {', '.join(available)})")
        selected = list(sheets)[0]

    rid = selected.attrib.get(f"{{{NS_REL_DOC}}}id")
    if not rid or rid not in rel_map:
        raise ValueError("ไม่พบความสัมพันธ์ของ Worksheet")

    target = rel_map[rid].replace("\\", "/")
    if target.startswith("/"):
        target = target.lstrip("/")
    elif not target.startswith("xl/"):
        target = str(PurePosixPath("xl") / target)

    # Normalize xl/../ style paths
    parts: list[str] = []
    for part in PurePosixPath(target).parts:
        if part == "..":
            if parts:
                parts.pop()
        elif part != ".":
            parts.append(part)
    normalized = "/".join(parts)
    return normalized, selected.attrib.get("name", "Sheet1")


def _parse_cell(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    value_node = cell.find(f"{{{NS_MAIN}}}v")

    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.iter(f"{{{NS_MAIN}}}t")]
        return "".join(texts)

    if value_node is None or value_node.text is None:
        return None

    raw = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "b":
        return raw == "1"
    if cell_type in {"str", "e"}:
        return raw

    try:
        number = float(raw)
        if number.is_integer():
            return int(number)
        return number
    except ValueError:
        return raw


def read_xlsx_rows(path: Path, sheet_name: str | None = None) -> tuple[list[list[Any]], str]:
    if not path.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("รองรับเฉพาะไฟล์ .xlsx")

    with zipfile.ZipFile(path) as zf:
        shared_strings = _read_shared_strings(zf)
        sheet_path, actual_sheet_name = _workbook_sheet_path(zf, sheet_name)
        if sheet_path not in zf.namelist():
            raise ValueError(f"ไม่พบไฟล์ Worksheet ภายใน xlsx: {sheet_path}")
        root = ET.fromstring(zf.read(sheet_path))

        output: list[list[Any]] = []
        sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
        if sheet_data is None:
            return [], actual_sheet_name

        for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
            values: dict[int, Any] = {}
            max_col = -1
            for cell in row.findall(f"{{{NS_MAIN}}}c"):
                ref = cell.attrib.get("r", "A1")
                col_idx = _column_index(ref)
                values[col_idx] = _parse_cell(cell, shared_strings)
                max_col = max(max_col, col_idx)
            if max_col >= 0:
                output.append([values.get(i) for i in range(max_col + 1)])
            else:
                output.append([])
    return output, actual_sheet_name


# -----------------------------
# Data preparation
# -----------------------------
FIELD_ALIASES: dict[str, list[str]] = {
    "company": ["บริษัทที่รับผิดชอบ", "บริษัท", "ผู้รับจ้าง"],
    "pole_code": ["รหัสเสาไฟฟ้า", "รหัสเสา", "lamp_post_code", "pole_code"],
    "lamp_type": ["ชนิดโคม", "ประเภทโคม", "lamp_type"],
    "lat": ["lat", "latitude", "ละติจูด"],
    "lon": ["lon", "lng", "longitude", "ลองจิจูด"],
    "damage_type": ["ประเภทความเสียหาย", "ประเภทเสียหาย", "damage_type"],
    "symptom": ["อาการที่ตรวจสอบ", "อาการ", "อาการเสีย", "symptom"],
    "operator": ["ผู้ดำเนินการ", "ผูดำเนินการ", "ผู้ปฏิบัติงาน", "operator"],
    "repair_method": ["วิธีแก้ไข", "วิธีการซ่อม", "วิธีซ่อม", "repair_method"],
    "details": ["รายละเอียดเพิ่มเติม", "รายละเอียดเพิ่มเติม ", "รายละเอียด", "detail"],
    "duration_text": ["ระยะเวลาดำเนินการ", "ช่วงเวลาดำเนินการ", "ระยะเวลา"],
    "ticket_id": ["ticket_id", "เลขที่ ticket", "เลขข้อร้องเรียน", "ticket"],
    "complaint_status": ["สถานะข้อร้องเรียน", "สถานะงาน", "status"],
    "survey_status": ["สถานะสำรวจ", "survey_status"],
    "district": ["เขต", "district", "zone", "โซน"],
    "agency": ["หน่วยงานรับผิดชอบ", "หน่วยงาน", "agency"],
    "contract_no": ["เลขที่สัญญา", "contract_no"],
    "contract_start": ["วันที่เริ่มต้นสัญญา", "contract_start"],
    "contract_end": ["วันที่สิ้นสุดสัญญา", "contract_end"],
    "warranty": ["ระยะค้ำประกันสัญญา", "ระยะค้ำประกัน", "warranty"],
    "post_repair_status": [
        "ผลหลังซ่อม", "สถานะหลังซ่อม", "สถานะโคมหลังซ่อม", "ผลการซ่อม", "post_repair_status"
    ],
}

PLACEHOLDER_POLE_CODES = {"", "_", "-", "0", "00", "000", "ไม่ระบุ", "n/a", "na", "null", "none"}
NO_REPAIR_METHODS = {
    "", "-", "ไม่ระบุ", "ไม่เสียหาย", "ไม่พบความเสียหาย", "ตรวจสอบแล้วปกติ", "ปกติ"
}


def clean_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"none", "null", "nan"}:
        return ""
    return re.sub(r"\s+", " ", text)


def parse_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def is_valid_coordinate(lat: float | None, lon: float | None) -> bool:
    return lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180


def is_bangkok_coordinate(lat: float | None, lon: float | None) -> bool:
    # กรอบกว้างครอบคลุมกรุงเทพฯ และพื้นที่รอยต่อ เพื่อใช้แจ้งเตือนเท่านั้น
    return bool(lat is not None and lon is not None and 13.35 <= lat <= 14.25 and 99.90 <= lon <= 101.05)


def parse_datetime_flexible(text: str) -> datetime | None:
    text = clean_text(text)
    if not text:
        return None
    formats = [
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
        "%d/%m/%Y", "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def parse_duration_range(text: str) -> tuple[datetime | None, datetime | None, float | None]:
    value = clean_text(text)
    if not value:
        return None, None, None

    # รองรับช่องว่างรอบเครื่องหมาย - และวันที่มี - อยู่ภายใน
    match = re.match(
        r"^\s*(\d{1,4}[/-]\d{1,2}[/-]\d{1,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)\s+-\s+"
        r"(\d{1,4}[/-]\d{1,2}[/-]\d{1,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)\s*$",
        value,
    )
    if not match:
        return None, None, None
    start = parse_datetime_flexible(match.group(1))
    end = parse_datetime_flexible(match.group(2))
    if not start or not end:
        return start, end, None
    hours = (end - start).total_seconds() / 3600
    if hours < 0:
        return start, end, None
    return start, end, round(hours, 2)


def contains_any(text: str, keywords: Iterable[str]) -> bool:
    lowered = clean_text(text).lower()
    return any(keyword.lower() in lowered for keyword in keywords)


def map_headers(headers: list[Any]) -> dict[str, int]:
    normalized = {clean_header(header): idx for idx, header in enumerate(headers)}
    mapped: dict[str, int] = {}
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            key = clean_header(alias)
            if key in normalized:
                mapped[field] = normalized[key]
                break
    return mapped


def row_value(row: list[Any], index: int | None) -> Any:
    if index is None or index >= len(row):
        return None
    return row[index]


BANGKOK_DISTRICTS = (
    "คลองเตย", "คลองสาน", "คลองสามวา", "คันนายาว", "จตุจักร", "จอมทอง", "ดอนเมือง", "ดินแดง",
    "ดุสิต", "ตลิ่งชัน", "ทวีวัฒนา", "ทุ่งครุ", "ธนบุรี", "บางกอกน้อย", "บางกอกใหญ่", "บางกะปิ",
    "บางขุนเทียน", "บางเขน", "บางคอแหลม", "บางแค", "บางซื่อ", "บางนา", "บางบอน", "บางพลัด",
    "บางรัก", "บึงกุ่ม", "ปทุมวัน", "ประเวศ", "ป้อมปราบศัตรูพ่าย", "พญาไท", "พระโขนง", "พระนคร",
    "ภาษีเจริญ", "มีนบุรี", "ยานนาวา", "ราชเทวี", "ราษฎร์บูรณะ", "ลาดกระบัง", "ลาดพร้าว", "วังทองหลาง",
    "วัฒนา", "สวนหลวง", "สะพานสูง", "สัมพันธวงศ์", "สาทร", "สายไหม", "หนองแขม", "หนองจอก",
    "หลักสี่", "ห้วยขวาง",
)
DISTRICT_BY_COMPACT_NAME = {re.sub(r"\s+", "", name): name for name in BANGKOK_DISTRICTS}


def normalize_bangkok_district(value: Any) -> str:
    name = clean_text(value)
    if name.startswith("เขต"):
        name = name[3:]
    return DISTRICT_BY_COMPACT_NAME.get(re.sub(r"\s+", "", name), "ไม่ระบุเขต")


def load_records(rows: list[list[Any]]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    nonempty_rows = [row for row in rows if any(clean_text(v) for v in row)]
    if not nonempty_rows:
        raise ValueError("Excel ไม่มีข้อมูล")

    headers = nonempty_rows[0]
    mapping = map_headers(headers)
    required = ["pole_code", "lat", "lon", "district"]
    missing = [field for field in required if field not in mapping]
    if missing:
        readable = {field: FIELD_ALIASES[field][0] for field in missing}
        raise ValueError("ไม่พบคอลัมน์ที่จำเป็น: " + ", ".join(readable.values()))

    has_post_repair_status = "post_repair_status" in mapping
    records: list[dict[str, Any]] = []
    invalid_coords = 0
    outside_bangkok = 0
    placeholder_codes = 0
    missing_methods_completed_outage = 0

    for source_row_no, row in enumerate(nonempty_rows[1:], start=2):
        if not any(clean_text(v) for v in row):
            continue

        record: dict[str, Any] = {field: clean_text(row_value(row, idx)) for field, idx in mapping.items()}
        for field in FIELD_ALIASES:
            record.setdefault(field, "")
        record["district"] = normalize_bangkok_district(record.get("district", ""))

        lat = parse_float(row_value(row, mapping.get("lat")))
        lon = parse_float(row_value(row, mapping.get("lon")))
        valid_coord = is_valid_coordinate(lat, lon)
        if not valid_coord:
            invalid_coords += 1
        elif not is_bangkok_coordinate(lat, lon):
            outside_bangkok += 1

        pole_code = clean_text(record["pole_code"])
        pole_placeholder = pole_code.lower() in PLACEHOLDER_POLE_CODES
        if pole_placeholder:
            placeholder_codes += 1

        symptom_all = " ".join([
            record.get("symptom", ""), record.get("details", ""), record.get("damage_type", "")
        ])
        is_outage = contains_any(symptom_all, ["ไฟดับ", "ไม่ติด", "ดับ", "ไม่มีแสง"])
        is_completed = contains_any(record.get("complaint_status", ""), ["เสร็จ", "สำเร็จ", "ปิดงาน", "แล้วเสร็จ"])

        method = clean_text(record.get("repair_method", ""))
        meaningful_method = method.lower() not in {v.lower() for v in NO_REPAIR_METHODS}
        post_status = clean_text(record.get("post_repair_status", ""))
        if has_post_repair_status and post_status:
            negative = contains_any(post_status, ["ไม่ติด", "ยังดับ", "ไม่สำเร็จ", "ใช้งานไม่ได้"])
            positive = contains_any(post_status, ["กลับมาติด", "ติด", "ปกติ", "ใช้งานได้", "สำเร็จ"])
            repair_success = is_outage and positive and not negative
            success_basis = "ผลหลังซ่อม"
        else:
            repair_success = is_outage and is_completed and meaningful_method
            success_basis = "สถานะเสร็จสิ้น (Proxy)"

        start_dt, end_dt, repair_hours = parse_duration_range(record.get("duration_text", ""))
        if is_outage and is_completed and not meaningful_method:
            missing_methods_completed_outage += 1

        # ใช้พิกัดช่วยสร้าง key เมื่อรหัสเสาเป็น placeholder
        if not pole_placeholder:
            pole_key = pole_code
        elif valid_coord:
            pole_key = f"coord:{lat:.6f},{lon:.6f}"
        else:
            pole_key = f"row:{source_row_no}"

        record.update({
            "source_row": source_row_no,
            "lat": lat,
            "lon": lon,
            "valid_coord": valid_coord,
            "outside_bangkok": valid_coord and not is_bangkok_coordinate(lat, lon),
            "pole_placeholder": pole_placeholder,
            "pole_key": pole_key,
            "is_outage": is_outage,
            "is_completed": is_completed,
            "repair_success": repair_success,
            "success_basis": success_basis,
            "repair_hours": repair_hours,
            "start_iso": start_dt.isoformat(timespec="seconds") if start_dt else "",
            "end_iso": end_dt.isoformat(timespec="seconds") if end_dt else "",
            "start_date": start_dt.strftime("%Y-%m-%d") if start_dt else "",
            "end_date": end_dt.strftime("%Y-%m-%d") if end_dt else "",
        })
        records.append(record)

    ticket_ids = [r["ticket_id"] for r in records if r.get("ticket_id")]
    duplicate_ticket_ids = sum(count - 1 for count in Counter(ticket_ids).values() if count > 1)
    missing_district = sum(1 for r in records if not r.get("district"))

    quality = {
        "input_rows": len(records),
        "valid_coordinate_rows": sum(1 for r in records if r["valid_coord"]),
        "invalid_coordinate_rows": invalid_coords,
        "outside_bangkok_rows": outside_bangkok,
        "placeholder_pole_codes": placeholder_codes,
        "duplicate_ticket_ids": duplicate_ticket_ids,
        "missing_district": missing_district,
        "missing_methods_completed_outage": missing_methods_completed_outage,
        "has_post_repair_status": has_post_repair_status,
        "success_definition": (
            "ใช้คอลัมน์ผลหลังซ่อม/สถานะหลังซ่อม" if has_post_repair_status
            else "ใช้ข้อร้องเรียนเสร็จสิ้นและมีวิธีแก้ไขเป็นตัวแทน (Proxy)"
        ),
    }
    return records, quality


def summarize_for_build(records: list[dict[str, Any]]) -> dict[str, Any]:
    valid = [r for r in records if r["valid_coord"]]
    lats = [r["lat"] for r in valid]
    lons = [r["lon"] for r in valid]
    if valid:
        center = [sum(lons) / len(lons), sum(lats) / len(lats)]
        bounds = [[min(lons), min(lats)], [max(lons), max(lats)]]
    else:
        center = [100.5018, 13.7563]
        bounds = [[100.30, 13.45], [100.95, 14.05]]

    dates = sorted(r["start_date"] for r in records if r.get("start_date"))
    return {
        "center": center,
        "bounds": bounds,
        "min_date": dates[0] if dates else "",
        "max_date": dates[-1] if dates else "",
    }


# -----------------------------
# HTML template
# -----------------------------
HTML_TEMPLATE = r'''<!doctype html>
<html lang="th">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta name="description" content="Dashboard วิเคราะห์การซ่อมเสาไฟด้วย MapLibre" />
  <title>Lighting Maintenance Intelligence</title>
  <link rel="preconnect" href="https://unpkg.com" />
  <link rel="stylesheet" href="https://unpkg.com/maplibre-gl@5.24.0/dist/maplibre-gl.css" />
  <script src="https://unpkg.com/maplibre-gl@5.24.0/dist/maplibre-gl.js"></script>
  <style>
    :root {
      --bg: #f3f6fb;
      --card: #ffffff;
      --ink: #132238;
      --muted: #64748b;
      --line: #dce4ee;
      --primary: #155eef;
      --primary-soft: #eaf1ff;
      --danger: #d92d20;
      --warning: #dc6803;
      --success: #039855;
      --shadow: 0 10px 28px rgba(15, 23, 42, 0.08);
      --radius: 16px;
    }
    * { box-sizing: border-box; }
    html, body { height: 100%; overflow: hidden; }
    body {
      margin: 0;
      font-family: "Noto Sans Thai", "Leelawadee UI", Tahoma, Arial, sans-serif;
      color: var(--ink);
      background: var(--bg);
    }
    button, input, select { font: inherit; }
    .app { height: 100%; min-height: 0; display: grid; grid-template-rows: auto minmax(0, 1fr); overflow: hidden; }
    header {
      min-height: 68px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      padding: 12px 22px;
      background: #102a43;
      color: white;
      box-shadow: 0 4px 14px rgba(15, 23, 42, 0.18);
      z-index: 10;
    }
    .brand { display: flex; align-items: center; gap: 13px; }
    .brand-icon {
      width: 42px; height: 42px; border-radius: 12px;
      display: grid; place-items: center; font-size: 22px;
      background: linear-gradient(135deg, #ffcf33, #ff8a00);
      box-shadow: 0 6px 16px rgba(255, 170, 0, .25);
    }
    .brand h1 { margin: 0; font-size: 19px; line-height: 1.2; }
    .brand p { margin: 4px 0 0; color: #cbd5e1; font-size: 12px; }
    .header-actions { display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
    .button {
      border: 1px solid rgba(255,255,255,.28);
      background: rgba(255,255,255,.10);
      color: white; border-radius: 10px; padding: 8px 12px;
      cursor: pointer; transition: .15s ease; font-weight: 650; font-size: 13px;
    }
    .button:hover { background: rgba(255,255,255,.18); transform: translateY(-1px); }
    .button.light { color: var(--ink); background: white; border-color: var(--line); }
    .button.primary { color: white; background: var(--primary); border-color: var(--primary); }
    .button.danger { color: var(--danger); background: #fff1f0; border-color: #fecdca; }

    .workspace {
      min-height: 0;
      display: grid;
      grid-template-columns: 294px minmax(520px, 1fr) 370px;
      gap: 14px;
      padding: 14px;
      overflow: hidden;
    }
    .workspace.filters-collapsed { grid-template-columns: minmax(520px, 1fr) 370px; }
    .workspace.filters-collapsed .sidebar { display: none; }
    .panel {
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      overflow: hidden;
      min-height: 0;
    }
    .sidebar { display: flex; flex-direction: column; height: 100%; }
    .panel-title {
      padding: 15px 16px 12px;
      border-bottom: 1px solid var(--line);
      display: flex; align-items: center; justify-content: space-between; gap: 8px;
    }
    .panel-title h2 { font-size: 15px; margin: 0; }
    .panel-title small { color: var(--muted); }
    .filters { flex: 1; min-height: 0; overflow-y: auto; overscroll-behavior: contain; scrollbar-gutter: stable; padding: 14px 15px 18px; }
    .field { margin-bottom: 12px; }
    .field label { display: block; font-size: 12px; font-weight: 700; margin-bottom: 6px; color: #344054; }
    .field select, .field input {
      width: 100%; border: 1px solid #cfd8e3; border-radius: 9px;
      background: white; padding: 9px 10px; outline: none; color: var(--ink);
    }
    .field select:focus, .field input:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(21,94,239,.12); }
    .date-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    .filter-buttons { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 16px; }
    .center-column { min-height: 0; display: grid; grid-template-rows: auto minmax(480px, 1fr); gap: 12px; }
    .kpi-row { display: grid; grid-template-columns: repeat(6, minmax(120px,1fr)); gap: 10px; }
    .kpi {
      background: white; border: 1px solid var(--line); border-radius: 13px;
      padding: 12px; box-shadow: 0 6px 20px rgba(15,23,42,.06);
      min-width: 0;
    }
    .kpi .label { font-size: 11px; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .kpi .value { font-size: 23px; font-weight: 800; margin-top: 6px; letter-spacing: -.5px; }
    .kpi .sub { font-size: 10px; margin-top: 3px; color: var(--muted); }
    .map-panel { position: relative; }
    #map { position: absolute; inset: 0; }
    .map-toolbar {
      position: absolute; left: 12px; top: 12px; right: 12px; z-index: 3;
      display: flex; justify-content: space-between; gap: 10px; pointer-events: none;
    }
    .mode-tabs {
      pointer-events: auto; display: flex; gap: 5px; padding: 5px;
      border-radius: 12px; background: rgba(255,255,255,.96); box-shadow: 0 8px 24px rgba(15,23,42,.16);
      overflow-x: auto;
    }
    .mode-button {
      border: 0; background: transparent; color: #475467; cursor: pointer;
      padding: 8px 11px; border-radius: 8px; font-size: 12px; font-weight: 700; white-space: nowrap;
    }
    .mode-button.active { background: #0b57d0; color: white; box-shadow: 0 2px 6px rgba(11,87,208,.28); }
    .layer-controls {
      position: absolute; top: 16px; right: 16px; z-index: 4;
      display: grid; gap: 10px; padding: 12px 14px;
      background: rgba(255,255,255,.97); border: 0;
      border-radius: 18px; box-shadow: 0 2px 4px rgba(60,64,67,.16), 0 8px 24px rgba(60,64,67,.14);
      font-size: 12px; font-weight: 700;
    }
    .layer-controls label { display: flex; align-items: center; gap: 9px; cursor: pointer; white-space: nowrap; }
    .layer-controls input { width: 17px; height: 17px; accent-color: #0b57d0; }
    .repair-filter-trigger {
      position: absolute; top: 50%; right: 0; z-index: 8; transform: translateY(-50%);
      display: flex; align-items: center; gap: 7px; min-height: 52px; max-width: 190px;
      border: 0; border-radius: 18px 0 0 18px; padding: 8px 12px 8px 9px;
      color: #0b57d0; background: #d3e3fd; cursor: pointer;
      box-shadow: -2px 3px 8px rgba(60,64,67,.22), -8px 8px 24px rgba(60,64,67,.13);
      font-size: 12px; font-weight: 800; transition: background .2s ease, color .2s ease, transform .2s ease;
    }
    .repair-filter-trigger:hover { background: #c2d7f7; transform: translateY(-50%) translateX(-3px); }
    .repair-filter-trigger.active { color: #fff; background: #0b57d0; }
    .filter-trigger-arrow { display: grid; place-items: center; width: 28px; height: 28px; flex: 0 0 28px; border-radius: 50%; background: rgba(255,255,255,.7); color: #0b57d0; font-size: 24px; line-height: 1; }
    .repair-filter-trigger.active .filter-trigger-arrow { background: rgba(255,255,255,.18); color: #fff; }
    .filter-trigger-label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .drawer-backdrop {
      position: absolute; inset: 0; z-index: 18; opacity: 0; visibility: hidden;
      background: rgba(32,33,36,.28); backdrop-filter: blur(1px); transition: opacity .3s ease, visibility .3s ease;
    }
    .drawer-backdrop.open { opacity: 1; visibility: visible; }
    .repair-drawer {
      position: absolute; top: 10px; right: 10px; bottom: 10px; z-index: 20; width: min(380px, calc(100vw - 28px));
      display: flex; flex-direction: column; overflow: hidden; background: #fefbff;
      border-radius: 28px 0 0 28px; box-shadow: -4px 0 12px rgba(60,64,67,.18), -18px 0 46px rgba(60,64,67,.16);
      transform: translateX(calc(100% + 18px)); transition: transform .34s cubic-bezier(.2,0,0,1);
    }
    .repair-drawer.open { transform: translateX(0); }
    .drawer-head { display: flex; align-items: center; gap: 12px; padding: 20px 18px 16px; background: #eef3fc; border-bottom: 1px solid #d7e3f6; }
    .drawer-title { flex: 1; min-width: 0; }
    .drawer-head h2 { margin: 0; color: #1f1f1f; font-size: 18px; font-weight: 700; }
    .drawer-head p { margin: 3px 0 0; color: #5f6368; font-size: 11px; }
    .drawer-close { width: 42px; height: 42px; border: 0; border-radius: 50%; padding: 0; color: #0b57d0; background: #d3e3fd; cursor: pointer; font-size: 26px; line-height: 1; transition: background .2s ease, transform .2s ease; }
    .drawer-close:hover { background: #c2d7f7; transform: translateX(2px); }
    .drawer-body { flex: 1; padding: 24px 20px; overflow: auto; }
    .drawer-body label { display: block; margin-bottom: 9px; color: #3c4043; font-size: 12px; font-weight: 700; }
    .material-select { position: relative; }
    .drawer-body select { width: 100%; min-height: 56px; appearance: none; padding: 14px 44px 14px 16px; border: 1px solid #747775; border-radius: 12px; outline: none; background: #fefbff; color: #1f1f1f; font-size: 14px; transition: border .2s ease, box-shadow .2s ease; }
    .drawer-body select:focus { border: 2px solid #0b57d0; box-shadow: 0 0 0 3px rgba(11,87,208,.10); }
    .select-arrow { position: absolute; top: 50%; right: 16px; transform: translateY(-50%); color: #444746; pointer-events: none; font-size: 17px; }
    .drawer-hint { margin-top: 12px; padding: 12px 14px; border-radius: 12px; color: #444746; background: #f0f4f9; font-size: 11px; line-height: 1.55; }
    .drawer-actions { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; padding: 16px 20px 20px; background: #fefbff; border-top: 1px solid #e1e3e1; }
    .drawer-actions button { min-height: 44px; border: 0; border-radius: 22px; padding: 10px 16px; cursor: pointer; font-weight: 800; transition: box-shadow .2s ease, transform .15s ease; }
    .drawer-actions button:hover { transform: translateY(-1px); }
    .drawer-actions .clear { color: #0b57d0; background: #e8f0fe; }
    .drawer-actions .apply { color: #fff; background: #0b57d0; box-shadow: 0 2px 6px rgba(11,87,208,.30); }
    .district-marker { display: grid; justify-items: center; cursor: pointer; font-family: "Noto Sans Thai", "Leelawadee UI", Tahoma, sans-serif; }
    .district-marker-name {
      margin-bottom: 3px; padding: 3px 7px; border-radius: 7px;
      color: #102a43; background: rgba(255,255,255,.96); border: 1px solid #b2ccff;
      box-shadow: 0 3px 9px rgba(15,23,42,.18); font-size: 11px; font-weight: 800;
      line-height: 1.2; white-space: nowrap;
    }
    .district-marker-count {
      min-width: 42px; height: 42px; padding: 0 7px; border-radius: 999px;
      display: grid; place-items: center; color: #fff; background: rgba(21,94,239,.88);
      border: 2px solid #fff; box-shadow: 0 4px 12px rgba(21,94,239,.30);
      font-size: 12px; font-weight: 800;
    }
    .map-mini-actions { pointer-events: auto; display: flex; gap: 6px; }
    .map-mini-actions .button { background: rgba(255,255,255,.96); color: var(--ink); border-color: var(--line); box-shadow: 0 8px 24px rgba(15,23,42,.12); }
    .map-legend {
      position: absolute; left: 12px; bottom: 30px; z-index: 3;
      background: rgba(255,255,255,.95); border: 1px solid var(--line); border-radius: 11px;
      padding: 10px 12px; max-width: 310px; box-shadow: 0 8px 24px rgba(15,23,42,.13);
      font-size: 11px;
    }
    .legend-title { font-weight: 800; margin-bottom: 7px; }
    .gradient { height: 8px; border-radius: 99px; background: linear-gradient(90deg,#2b83ba,#abdda4,#ffffbf,#fdae61,#d7191c); }
    .legend-labels { display: flex; justify-content: space-between; color: var(--muted); margin-top: 4px; }
    .map-status {
      position: absolute; right: 12px; bottom: 30px; z-index: 3;
      background: rgba(16,42,67,.92); color: white; padding: 7px 10px; border-radius: 9px; font-size: 11px;
    }

    .analytics { display: block; height: 100%; overflow-y: auto; overscroll-behavior: contain; scrollbar-gutter: stable; }
    .analytics-section { padding: 14px 15px; border-bottom: 1px solid var(--line); }
    .analytics-section h3 { margin: 0 0 10px; font-size: 14px; }
    .analytics-section .hint { color: var(--muted); font-size: 10px; margin: -5px 0 9px; }
    .bar-list { display: grid; gap: 8px; }
    .bar-row { cursor: pointer; }
    .bar-meta { display: flex; justify-content: space-between; gap: 10px; font-size: 11px; margin-bottom: 3px; }
    .bar-name { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
    .bar-value { font-weight: 800; }
    .bar-track { height: 8px; border-radius: 99px; background: #edf1f6; overflow: hidden; }
    .bar-fill { height: 100%; border-radius: 99px; background: linear-gradient(90deg,#155eef,#53b1fd); min-width: 2px; }
    .bar-row.method .bar-fill { background: linear-gradient(90deg,#039855,#6ce9a6); }
    .table-wrap { overflow: auto; max-height: 420px; border: 1px solid var(--line); border-radius: 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 10.5px; }
    thead { position: sticky; top: 0; background: #f8fafc; z-index: 1; }
    th, td { padding: 8px 7px; border-bottom: 1px solid #eef2f6; text-align: left; vertical-align: top; }
    th { color: #475467; font-weight: 800; white-space: nowrap; }
    tbody tr { cursor: pointer; }
    tbody tr:hover { background: var(--primary-soft); }
    .pill { display: inline-flex; align-items: center; border-radius: 99px; padding: 2px 7px; font-size: 9.5px; font-weight: 700; }
    .pill.red { color: #b42318; background: #fee4e2; }
    .pill.green { color: #027a48; background: #d1fadf; }
    .pill.orange { color: #b54708; background: #fef0c7; }
    .empty { color: var(--muted); text-align: center; padding: 18px; }
    .pole-detail-empty { color: var(--muted); text-align: center; padding: 28px 12px; line-height: 1.7; }
    .pole-detail-head { padding-bottom: 11px; border-bottom: 1px solid var(--line); }
    .pole-detail-head strong { display: block; font-size: 14px; word-break: break-word; }
    .pole-detail-head small { display: block; color: var(--muted); margin-top: 4px; }
    .detail-stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 7px; margin: 11px 0; }
    .detail-stat { padding: 8px 6px; border-radius: 9px; background: #f8fafc; text-align: center; }
    .detail-stat strong { display: block; font-size: 16px; }
    .detail-stat span { color: var(--muted); font-size: 9px; }
    .detail-meta { display: grid; grid-template-columns: 82px 1fr; gap: 5px 8px; font-size: 10.5px; margin-bottom: 12px; }
    .detail-meta span:nth-child(odd) { color: var(--muted); }
    .detail-meta span:nth-child(even) { font-weight: 650; word-break: break-word; }
    .history-title { font-size: 12px; font-weight: 800; margin: 13px 0 7px; }
    .history-list { display: grid; gap: 8px; max-height: 360px; overflow: auto; }
    .history-item { border: 1px solid var(--line); border-radius: 9px; padding: 9px; font-size: 10px; line-height: 1.5; }
    .history-item-head { display: flex; justify-content: space-between; gap: 8px; margin-bottom: 5px; }
    .history-item-head strong { font-size: 10.5px; }
    .history-item-head span { color: var(--muted); white-space: nowrap; }
    .history-row { display: grid; grid-template-columns: 70px 1fr; gap: 5px; }
    .history-row span:first-child { color: var(--muted); }
    .history-row span:last-child { font-weight: 650; word-break: break-word; }

    .maplibregl-popup-content { border-radius: 12px; padding: 0; overflow: hidden; box-shadow: 0 12px 35px rgba(15,23,42,.22); }
    .popup { min-width: 245px; max-width: 330px; }
    .popup-head { background: #102a43; color: white; padding: 11px 13px; }
    .popup-head strong { font-size: 13px; }
    .popup-head small { display:block; margin-top:3px; color:#cbd5e1; }
    .popup-body { padding: 10px 13px 12px; font-size: 11px; }
    .popup-grid { display:grid; grid-template-columns: 96px 1fr; gap:5px 8px; }
    .popup-grid span:nth-child(odd) { color: var(--muted); }
    .popup-grid span:nth-child(even) { font-weight: 650; word-break: break-word; }

    /* Full-screen heatmap view: retain only the map and its mode switcher. */
    header, .sidebar, .analytics, .kpi-row, .map-mini-actions, .map-legend, .map-status { display: none !important; }
    .app { display: block; height: 100%; }
    .workspace, .workspace.filters-collapsed { display: block; height: 100%; padding: 0; overflow: hidden; }
    .center-column { display: block; height: 100%; min-height: 0; }
    .map-panel { height: 100%; border: 0; border-radius: 0; box-shadow: none; }
    .map-toolbar { left: 16px; top: 16px; right: auto; }
    .mode-tabs { max-width: calc(100vw - 32px); }

    @media (max-width: 1280px) {
      .workspace { grid-template-columns: 250px minmax(480px, 1fr) 320px; gap: 10px; padding: 10px; }
      .kpi-row { grid-template-columns: repeat(3, 1fr); }
    }
    @media (max-width: 1050px) {
      html, body { height: 100%; overflow: hidden; }
      .app { height: 100%; min-height: 0; overflow: hidden; }
      header { align-items: flex-start; flex-direction: column; }
      .workspace, .workspace.filters-collapsed { display: block; height: 100%; padding: 0; overflow: hidden; }
      .sidebar, .analytics { height: auto; max-height: none; overflow: visible; }
      .filters { overflow: visible; }
      .center-column { height: 100%; min-height: 0; }
      .map-panel { height: 100%; min-height: 0; }
      .kpi-row { grid-template-columns: repeat(2, 1fr); }
      .analytics { display: block; }
      .analytics-section { border-right:0; border-bottom:1px solid var(--line); }
      .map-toolbar { flex-direction: column; align-items:flex-start; }
      .map-mini-actions { align-self:flex-end; }
      .map-status { left:12px; right:auto; bottom:12px; }
      .map-legend { bottom:48px; }
    }
  </style>
</head>
<body>
<div class="app">
  <header>
    <div class="brand">
      <div class="brand-icon">💡</div>
      <div>
        <h1>Lighting Maintenance Intelligence</h1>
        <p>วิเคราะห์เหตุขัดข้อง ประสิทธิผลการซ่อม และพื้นที่เสี่ยง</p>
      </div>
    </div>
    <div class="header-actions">
      <button class="button" id="filterToggleButton" aria-expanded="true" aria-controls="filterSidebar">ปิดตัวกรอง</button>
      <button class="button" id="exportButton">ส่งออก CSV ที่กรอง</button>
    </div>
  </header>

  <main class="workspace">
    <aside class="panel sidebar" id="filterSidebar">
      <div class="panel-title"><h2>ตัวกรองข้อมูล</h2><small id="filterCount">0 รายการ</small></div>
      <div class="filters">
        <div class="field">
          <label for="searchInput">ค้นหารหัสเสา / Ticket / รายละเอียด</label>
          <input id="searchInput" type="search" placeholder="เช่น lamp-001 หรือ 2026-..." />
        </div>
        <div class="field"><label for="districtFilter">เขต / Zone</label><select id="districtFilter"></select></div>
        <div class="field"><label for="lampFilter">ชนิดโคม</label><select id="lampFilter"></select></div>
        <div class="field"><label for="damageFilter">ประเภทความเสียหาย</label><select id="damageFilter"></select></div>
        <div class="field"><label for="symptomFilter">อาการที่ตรวจสอบ</label><select id="symptomFilter"></select></div>
        <div class="field"><label for="methodFilter">วิธีแก้ไข</label><select id="methodFilter"></select></div>
        <div class="field"><label for="statusFilter">สถานะข้อร้องเรียน</label><select id="statusFilter"></select></div>
        <div class="field">
          <label>วันที่เริ่มดำเนินการ</label>
          <div class="date-grid">
            <input id="dateFrom" type="date" value="__MIN_DATE__" />
            <input id="dateTo" type="date" value="__MAX_DATE__" />
          </div>
        </div>
        <div class="filter-buttons">
          <button class="button primary" id="applyButton">ใช้ตัวกรอง</button>
          <button class="button danger" id="resetButton">ล้างตัวกรอง</button>
        </div>
      </div>
    </aside>

    <section class="center-column">
      <div class="kpi-row">
        <div class="kpi"><div class="label">ข้อร้องเรียน</div><div class="value" id="kpiTickets">0</div><div class="sub">นับ Ticket ID ไม่ซ้ำ</div></div>
        <div class="kpi"><div class="label">เสาไฟไม่ซ้ำ</div><div class="value" id="kpiPoles">0</div><div class="sub">ยึดรหัสเสาไฟเป็นหลัก</div></div>
        <div class="kpi"><div class="label">ครั้งที่ไฟดับ</div><div class="value" id="kpiOutages">0</div><div class="sub">รวมข้อร้องเรียนในรอบซ่อมเดียวกัน</div></div>
        <div class="kpi"><div class="label">อัตราปิดงานซ่อม</div><div class="value" id="kpiCompleted">0%</div><div class="sub">คำนวณจากรอบการซ่อม</div></div>
        <div class="kpi"><div class="label">เวลาซ่อมค่ากลาง</div><div class="value" id="kpiMedian">–</div><div class="sub">Median ชั่วโมง</div></div>
        <div class="kpi"><div class="label">เสาซ่อมซ้ำ</div><div class="value" id="kpiRepeat">0</div><div class="sub">มีมากกว่า 1 รอบซ่อม</div></div>
      </div>

      <div class="panel map-panel">
        <div id="map"></div>
        <div class="map-toolbar">
          <div class="mode-tabs" role="tablist">
            <button class="mode-button active" data-mode="complaintHeat">ข้อร้องเรียน Heatmap</button>
            <button class="mode-button" data-mode="outagePoints">ไฟดับ จำนวนครั้ง</button>
            <button class="mode-button" data-mode="outageHeat">ไฟดับ Heatmap</button>
            <button class="mode-button" data-mode="repairPoints">การซ่อม จำนวนครั้ง</button>
            <button class="mode-button" data-mode="repairHeat">การซ่อม Heatmap</button>
          </div>
          <div class="map-mini-actions">
            <button class="button light" id="fitButton">แสดงทั้งหมด</button>
          </div>
        </div>
        <div class="map-legend" id="legend"></div>
        <div class="map-status" id="mapStatus">กำลังโหลดแผนที่…</div>
        <div class="layer-controls">
          <label><input id="heatmapToggle" type="checkbox" checked /> แสดง Heatmap</label>
          <label><input id="poleLocationToggle" type="checkbox" /> แสดงตำแหน่งเสา</label>
        </div>
        <button class="repair-filter-trigger" id="repairFilterOpen" aria-controls="repairFilterDrawer" aria-expanded="false">
          <span class="filter-trigger-arrow" aria-hidden="true">‹</span>
          <span class="filter-trigger-label" id="repairFilterOpenLabel">วิธีการแก้ไข</span>
        </button>
        <div class="drawer-backdrop" id="repairFilterBackdrop"></div>
        <aside class="repair-drawer" id="repairFilterDrawer" aria-hidden="true">
          <div class="drawer-head">
            <div class="drawer-title"><h2>วิธีการแก้ไข</h2><p>กรองข้อมูลบนแผนที่ทุกโหมด</p></div>
            <button class="drawer-close" id="repairFilterClose" type="button" aria-label="ปิดตัวกรอง">›</button>
          </div>
          <div class="drawer-body">
            <label for="repairMethodDrawerSelect">เลือกวิธีการแก้ไข</label>
            <div class="material-select"><select id="repairMethodDrawerSelect"></select><span class="select-arrow">▼</span></div>
            <div class="drawer-hint">เมื่อกดใช้ตัวกรอง แผนที่ Heatmap จุดเสา และยอดรวมรายเขตจะคำนวณใหม่ตามวิธีที่เลือก</div>
          </div>
          <div class="drawer-actions">
            <button class="clear" id="repairFilterClear" type="button">ล้างตัวกรอง</button>
            <button class="apply" id="repairFilterApply" type="button">ใช้ตัวกรอง</button>
          </div>
        </aside>
      </div>
    </section>

    <aside class="panel analytics">
      <section class="analytics-section">
        <h3>ข้อมูลรายละเอียดเสา</h3>
        <div class="hint">คลิกจุดเสาในแผนที่เพื่อดูประวัติการซ่อม</div>
        <div id="poleDetail" class="pole-detail-empty">ยังไม่ได้เลือกเสา<br />เลือกจุดบนแผนที่เพื่อดูสาเหตุ วิธีซ่อม และจำนวนข้อร้องเรียน</div>
      </section>
      <section class="analytics-section">
        <h3>เขตที่มีงานซ่อมสูงสุด</h3>
        <div class="hint">คลิกชื่อเขตเพื่อกรองแผนที่</div>
        <div class="bar-list" id="districtBars"></div>
      </section>
      <section class="analytics-section">
        <h3>วิธีซ่อมที่บันทึกไว้</h3>
        <div class="hint">นับหนึ่งครั้งต่อรอบการซ่อม ไม่ได้นับตาม Ticket</div>
        <div class="bar-list" id="methodBars"></div>
      </section>
    </aside>
  </main>
</div>

<script>
const RAW_RECORDS = __RECORDS_JSON__;
const INITIAL_CENTER = __CENTER_JSON__;
const INITIAL_BOUNDS = __BOUNDS_JSON__;

const ALL_VALUE = "__ALL__";
const UNKNOWN = "ไม่ระบุ";
const POLE_DETAIL_ZOOM = 15;
let filteredRecords = [...RAW_RECORDS];
let currentMode = "complaintHeat";
let mapReady = false;
let poleIndex = new Map();
let selectedPoleKey = "";
let districtMarkers = [];
let districtMarkerMetric = "repair_count";
let districtMarkersEnabled = false;

const el = id => document.getElementById(id);
const clean = v => String(v ?? "").trim();
const display = v => clean(v) || UNKNOWN;
const fmt = new Intl.NumberFormat("th-TH");
const fmt1 = new Intl.NumberFormat("th-TH", { maximumFractionDigits: 1 });
const escapeHtml = value => String(value ?? "")
  .replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
  .replaceAll('"', "&quot;").replaceAll("'", "&#039;");

function uniqueValues(field) {
  return [...new Set(RAW_RECORDS.map(r => display(r[field])))].sort((a,b) => a.localeCompare(b, "th"));
}

function fillSelect(id, field) {
  const select = el(id);
  select.innerHTML = `<option value="${ALL_VALUE}">ทั้งหมด</option>` +
    uniqueValues(field).map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
}

function fillRepairMethodDrawer() {
  const counts = new Map();
  RAW_RECORDS.forEach(record => {
    const method = display(record.repair_method);
    counts.set(method, (counts.get(method) || 0) + 1);
  });
  const options = [...counts.entries()].sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0], "th"));
  el("repairMethodDrawerSelect").innerHTML = `<option value="${ALL_VALUE}">ทั้งหมด</option>` +
    options.map(([method, count]) => `<option value="${escapeHtml(method)}">${escapeHtml(method)} (${fmt.format(count)})</option>`).join("");
}

function setRepairDrawerOpen(open) {
  el("repairFilterDrawer").classList.toggle("open", open);
  el("repairFilterBackdrop").classList.toggle("open", open);
  el("repairFilterDrawer").setAttribute("aria-hidden", String(!open));
  el("repairFilterOpen").setAttribute("aria-expanded", String(open));
}

function syncRepairFilterButton() {
  const value = selected("methodFilter");
  const active = value !== ALL_VALUE;
  el("repairFilterOpen").classList.toggle("active", active);
  el("repairFilterOpenLabel").textContent = active ? value : "วิธีการแก้ไข";
  el("repairFilterOpen").title = active ? value : "เปิดตัวกรองวิธีการแก้ไข";
}

function selected(id) { return el(id).value; }
function setSelected(id, value) {
  const select = el(id);
  const exists = [...select.options].some(o => o.value === value);
  select.value = exists ? value : ALL_VALUE;
}

function getFilteredRecords() {
  const query = clean(el("searchInput").value).toLowerCase();
  const from = el("dateFrom").value;
  const to = el("dateTo").value;
  const constraints = [
    ["district", selected("districtFilter")],
    ["lamp_type", selected("lampFilter")],
    ["damage_type", selected("damageFilter")],
    ["symptom", selected("symptomFilter")],
    ["repair_method", selected("methodFilter")],
    ["complaint_status", selected("statusFilter")],
  ];

  return RAW_RECORDS.filter(r => {
    for (const [field, expected] of constraints) {
      if (expected !== ALL_VALUE && display(r[field]) !== expected) return false;
    }
    if (from && r.start_date && r.start_date < from) return false;
    if (to && r.start_date && r.start_date > to) return false;
    if (query) {
      const haystack = [r.pole_code, r.ticket_id, r.details, r.symptom, r.repair_method, r.district, r.operator]
        .map(clean).join(" ").toLowerCase();
      if (!haystack.includes(query)) return false;
    }
    return true;
  });
}

function median(values) {
  const nums = values.filter(Number.isFinite).sort((a,b) => a-b);
  if (!nums.length) return null;
  const mid = Math.floor(nums.length / 2);
  return nums.length % 2 ? nums[mid] : (nums[mid-1] + nums[mid]) / 2;
}

function modeColorValue(mode, pole) {
  if (mode === "outage") return pole.outage_count;
  if (mode === "success") return pole.success_count;
  if (mode === "repeat") return pole.outage_count;
  return pole.repair_count;
}

function repairEventKey(r) {
  const start = clean(r.start_iso);
  const end = clean(r.end_iso);
  if (start || end) return `time:${start}|${end}`;
  const duration = clean(r.duration_text);
  if (duration) return `duration:${duration}`;
  const fallback = [r.start_date, r.repair_method, r.damage_type, r.symptom].map(clean).join("|");
  return fallback.replaceAll("|", "") ? `detail:${fallback}` : `row:${r.source_row}`;
}

function complaintCount(records) {
  return new Set(records.map(r => clean(r.ticket_id)).filter(Boolean)).size;
}

function mostCommon(records, field) {
  const counts = new Map();
  records.forEach(r => {
    const value = display(r[field]);
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return [...counts.entries()].sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0], "th"))[0]?.[0] || UNKNOWN;
}

function aggregateRepairEvents(records) {
  const groups = new Map();
  records.forEach(r => {
    const key = repairEventKey(r);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });
  return [...groups.entries()].map(([eventKey, items]) => {
    const first = items[0];
    const ticketIds = [...new Set(items.map(r => clean(r.ticket_id)).filter(Boolean))];
    const details = [...new Set(items.map(r => clean(r.details)).filter(Boolean))];
    const symptoms = [...new Set(items.map(r => clean(r.symptom)).filter(Boolean))];
    return {
      event_key: eventKey,
      start_iso: first.start_iso || "",
      end_iso: first.end_iso || "",
      duration_text: first.duration_text || "",
      repair_hours: median(items.map(r => r.repair_hours)),
      is_outage: items.some(r => r.is_outage),
      is_completed: items.some(r => r.is_completed),
      repair_success: items.some(r => r.repair_success),
      damage_type: mostCommon(items, "damage_type"),
      repair_method: mostCommon(items, "repair_method"),
      symptom: symptoms.join(", ") || UNKNOWN,
      details: details.join(", ") || UNKNOWN,
      complaint_count: ticketIds.length,
      ticket_ids: ticketIds.join(", "),
    };
  }).sort((a,b) => clean(b.start_iso).localeCompare(clean(a.start_iso)));
}

function aggregatePoles(records) {
  const groups = new Map();
  for (const r of records) {
    if (!r.valid_coord) continue;
    if (!groups.has(r.pole_key)) groups.set(r.pole_key, []);
    groups.get(r.pole_key).push(r);
  }
  const poles = [];
  for (const [poleKey, items] of groups.entries()) {
    const first = items[0];
    const incidents = aggregateRepairEvents(items);
    const successful = incidents.filter(r => display(r.repair_method) !== UNKNOWN);
    const methods = new Map();
    for (const r of successful) {
      const method = display(r.repair_method);
      methods.set(method, (methods.get(method) || 0) + 1);
    }
    const sortedMethods = [...methods.entries()].sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0], "th"));
    const topMethod = sortedMethods[0]?.[0] || UNKNOWN;
    const topMethodCount = sortedMethods[0]?.[1] || 0;
    const completed = incidents.filter(r => r.is_completed).length;
    const outages = incidents.filter(r => r.is_outage).length;
    const successCount = incidents.filter(r => r.repair_success).length;
    const hours = incidents.map(r => r.repair_hours).filter(Number.isFinite);
    const districts = new Map();
    items.forEach(r => districts.set(display(r.district), (districts.get(display(r.district)) || 0) + 1));
    const district = [...districts.entries()].sort((a,b)=>b[1]-a[1])[0]?.[0] || UNKNOWN;
    poles.push({
      pole_key: poleKey,
      pole_code: display(first.pole_code),
      lat: first.lat,
      lon: first.lon,
      district,
      lamp_type: display(first.lamp_type),
      complaint_count: complaintCount(items),
      ticket_count: complaintCount(items),
      repair_count: incidents.length,
      outage_count: outages,
      completed_count: completed,
      success_count: successCount,
      success_rate: outages ? successCount / outages : 0,
      median_hours: median(hours),
      top_method: topMethod,
      top_method_count: topMethodCount,
      damage_type: mostCommon(items, "damage_type"),
      incidents,
      ticket_ids: [...new Set(items.map(r => r.ticket_id).filter(Boolean))].slice(0, 12).join(", "),
      latest_detail: [...items].reverse().find(r => clean(r.details))?.details || "",
    });
  }
  return poles;
}

function eventGeoJSON(records) {
  return {
    type: "FeatureCollection",
    features: records.filter(r => r.valid_coord).map((r, i) => ({
      type: "Feature",
      id: i,
      geometry: { type: "Point", coordinates: [r.lon, r.lat] },
      properties: {
        pole_key: r.pole_key, pole_code: display(r.pole_code), ticket_id: display(r.ticket_id), district: display(r.district),
        lamp_type: display(r.lamp_type), damage_type: display(r.damage_type), symptom: display(r.symptom),
        repair_method: display(r.repair_method), complaint_status: display(r.complaint_status),
        details: display(r.details), duration_text: display(r.duration_text), repair_hours: r.repair_hours ?? -1,
        is_outage: r.is_outage ? 1 : 0, repair_success: r.repair_success ? 1 : 0,
      }
    }))
  };
}

function poleGeoJSON(poles) {
  return {
    type: "FeatureCollection",
    features: poles.map((p, i) => ({
      type: "Feature", id: i,
      geometry: { type: "Point", coordinates: [p.lon, p.lat] },
      properties: {
        pole_key:p.pole_key, pole_code:p.pole_code, district:p.district, lamp_type:p.lamp_type,
        complaint_count:p.complaint_count, repair_count:p.repair_count, outage_count:p.outage_count,
        completed_count:p.completed_count, success_count:p.success_count, top_method:p.top_method,
        median_hours:p.median_hours ?? -1, ticket_ids:p.ticket_ids, latest_detail:p.latest_detail
      }
    }))
  };
}

function districtGeoJSON(poles) {
  const districts = new Map();
  poles.forEach(p => {
    const name = display(p.district);
    if (!districts.has(name)) districts.set(name, {district:name, latSum:0, lonSum:0, poles:0, complaint_count:0, outage_count:0, repair_count:0});
    const d = districts.get(name);
    d.latSum += p.lat; d.lonSum += p.lon; d.poles += 1;
    d.complaint_count += p.complaint_count; d.outage_count += p.outage_count; d.repair_count += p.repair_count;
  });
  return {
    type:"FeatureCollection",
    features:[...districts.values()].map((d, i) => ({
      type:"Feature", id:i,
      geometry:{type:"Point", coordinates:[d.lonSum/d.poles, d.latSum/d.poles]},
      properties:{district:d.district, pole_count:d.poles, complaint_count:d.complaint_count, outage_count:d.outage_count, repair_count:d.repair_count}
    }))
  };
}

function rebuildDistrictMarkers(poles) {
  districtMarkers.forEach(item => item.marker.remove());
  districtMarkers = districtGeoJSON(poles).features.map(feature => {
    const node = document.createElement("div");
    node.className = "district-marker";
    const name = document.createElement("div");
    name.className = "district-marker-name";
    const value = document.createElement("div");
    value.className = "district-marker-count";
    name.textContent = feature.properties.district;
    node.append(name, value);
    node.addEventListener("click", event => {
      event.stopPropagation();
      map.easeTo({center:feature.geometry.coordinates, zoom:POLE_DETAIL_ZOOM + 0.5, duration:650});
    });
    const marker = new maplibregl.Marker({element:node, anchor:"bottom"})
      .setLngLat(feature.geometry.coordinates).addTo(map);
    return {marker, node, value, properties:feature.properties};
  });
  updateDistrictMarkerVisibility();
}

function updateDistrictMarkerVisibility() {
  const visible = districtMarkersEnabled && map.getZoom() < POLE_DETAIL_ZOOM;
  districtMarkers.forEach(item => {
    const count = Number(item.properties[districtMarkerMetric]) || 0;
    item.value.textContent = fmt.format(count);
    item.node.style.display = visible && count > 0 ? "grid" : "none";
  });
}

function countBy(records, field, predicate = () => true) {
  const counts = new Map();
  records.filter(predicate).forEach(r => {
    const key = display(r[field]);
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return [...counts.entries()].sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0], "th"));
}

function updateKpis(records, poles) {
  const incidents = poles.flatMap(p => p.incidents);
  const completed = incidents.filter(r => r.is_completed).length;
  const outage = incidents.filter(r => r.is_outage).length;
  const repairMedian = median(incidents.map(r => r.repair_hours));
  const repeatPoles = poles.filter(p => p.repair_count > 1).length;
  el("kpiTickets").textContent = fmt.format(complaintCount(records));
  el("kpiPoles").textContent = fmt.format(poles.length);
  el("kpiOutages").textContent = fmt.format(outage);
  el("kpiCompleted").textContent = incidents.length ? `${fmt1.format(completed / incidents.length * 100)}%` : "0%";
  el("kpiMedian").textContent = repairMedian == null ? "–" : `${fmt1.format(repairMedian)} ชม.`;
  el("kpiRepeat").textContent = fmt.format(repeatPoles);
  el("filterCount").textContent = `${fmt.format(complaintCount(records))} ข้อร้องเรียน`;
}

function renderBars(containerId, entries, cssClass, onClick) {
  const container = el(containerId);
  const top = entries.slice(0, 8);
  if (!top.length) {
    container.innerHTML = `<div class="empty">ไม่มีข้อมูลตามตัวกรอง</div>`;
    return;
  }
  const max = Math.max(...top.map(x => x[1]), 1);
  container.innerHTML = top.map(([name, value]) => `
    <div class="bar-row ${cssClass}" data-name="${escapeHtml(name)}" title="คลิกเพื่อกรอง ${escapeHtml(name)}">
      <div class="bar-meta"><span class="bar-name">${escapeHtml(name)}</span><span class="bar-value">${fmt.format(value)}</span></div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(3, value/max*100)}%"></div></div>
    </div>`).join("");
  container.querySelectorAll(".bar-row").forEach(row => row.addEventListener("click", () => onClick(row.dataset.name)));
}

function setLayerVisibility(id, visible) {
  if (map.getLayer(id)) map.setLayoutProperty(id, "visibility", visible ? "visible" : "none");
}

function modeFilter(mode) {
  if (mode === "outage") return [">", ["get", "outage_count"], 0];
  if (mode === "success") return [">", ["get", "repair_count"], 0];
  if (mode === "repeat") return [">", ["get", "repair_count"], 1];
  return [">=", ["get", "repair_count"], 0];
}

function configureMode(mode) {
  currentMode = mode;
  document.querySelectorAll(".mode-button").forEach(btn => btn.classList.toggle("active", btn.dataset.mode === mode));
  if (!mapReady) return;

  const modes = {
    complaintHeat: {property:"complaint_count", pointMode:false},
    outagePoints: {property:"outage_count", pointMode:true},
    outageHeat: {property:"outage_count", pointMode:false},
    repairPoints: {property:"repair_count", pointMode:true},
    repairHeat: {property:"repair_count", pointMode:false}
  };
  const config = modes[mode] || modes.complaintHeat;
  const filter = [">",["get",config.property],0];
  el("heatmapToggle").checked = !config.pointMode;
  setLayerVisibility("repair-heat", !config.pointMode);
  setLayerVisibility("event-points", false);
  setLayerVisibility("district-points", false);
  setLayerVisibility("district-labels", false);
  setLayerVisibility("pole-points", config.pointMode);
  setLayerVisibility("pole-labels", config.pointMode);
  setLayerVisibility("pole-locations", el("poleLocationToggle").checked && !config.pointMode);
  map.setFilter("repair-heat", filter);
  map.setPaintProperty("repair-heat", "heatmap-weight", [
    "interpolate", ["linear"], ["get", config.property], 0, 0, 1, 0.25, 3, 0.6, 8, 1
  ]);
  districtMarkerMetric = config.property;
  districtMarkersEnabled = config.pointMode;
  updateDistrictMarkerVisibility();
  if (config.pointMode) {
    ["pole-points","pole-labels"].forEach(id => map.setFilter(id, filter));
    map.setLayoutProperty("pole-labels", "text-field", ["to-string", ["get", config.property]]);
    map.setPaintProperty("pole-points", "circle-radius", [
      "interpolate", ["linear"], ["get", config.property], 1, 8, 3, 12, 8, 18
    ]);
    map.setPaintProperty("pole-points", "circle-color", [
      "interpolate", ["linear"], ["get", config.property], 1, "#53b1fd", 3, "#155eef", 8, "#102a43"
    ]);
    map.setPaintProperty("pole-points", "circle-opacity", 0.82);
    map.setPaintProperty("pole-points", "circle-stroke-color", "#ffffff");
    map.setPaintProperty("pole-points", "circle-stroke-width", 1.5);
  }
}

function refreshLayerToggles() {
  const pointMode = currentMode === "outagePoints" || currentMode === "repairPoints";
  setLayerVisibility("repair-heat", el("heatmapToggle").checked);
  setLayerVisibility("pole-locations", el("poleLocationToggle").checked && !pointMode);
}

function updateLegend(mode) {
  const legend = el("legend");
  if (mode === "heat") {
    legend.innerHTML = `<div class="legend-title">ความหนาแน่นของเสาซ่อมซ้ำ</div><div class="gradient"></div><div class="legend-labels"><span>น้อย</span><span>มาก</span></div><div style="margin-top:6px">แสดงเฉพาะเสาที่ซ่อมมากกว่า 1 รอบ</div>`;
  } else if (mode === "outage") {
    legend.innerHTML = `<div class="legend-title">จำนวนครั้งไฟดับต่อรหัสเสา</div><div>ข้อร้องเรียนช่วงซ่อมเดียวกันนับเป็นหนึ่งครั้ง</div>`;
  } else if (mode === "success") {
    legend.innerHTML = `<div class="legend-title">จำนวนรอบที่มีการซ่อม</div><div>คลิกจุดเพื่อดูวิธีซ่อมและประวัติของเสา</div>`;
  } else {
    legend.innerHTML = `<div class="legend-title">เสาที่ซ่อมซ้ำ</div><div>แสดงเฉพาะรหัสเสาที่มีมากกว่า 1 รอบซ่อม</div>`;
  }
}

function showPolePopup(feature, lngLat) {
  const p = feature.properties;
  const hours = Number(p.median_hours);
  const medianText = Number.isFinite(hours) && hours >= 0 ? `${fmt1.format(hours)} ชม.` : UNKNOWN;
  new maplibregl.Popup({closeButton:true, maxWidth:"360px"})
    .setLngLat(lngLat)
    .setHTML(`<div class="popup">
      <div class="popup-head"><strong>${escapeHtml(display(p.pole_code))}</strong><small>${escapeHtml(display(p.district))} · ${escapeHtml(display(p.lamp_type))}</small></div>
      <div class="popup-body"><div class="popup-grid">
        <span>จำนวนการซ่อม</span><span>${fmt.format(Number(p.repair_count) || 0)} รอบ</span>
        <span>จำนวนไฟดับ</span><span>${fmt.format(Number(p.outage_count) || 0)} ครั้ง</span>
        <span>ข้อร้องเรียน</span><span>${fmt.format(Number(p.complaint_count) || 0)} รายการ</span>
        <span>วิธีซ่อมหลัก</span><span>${escapeHtml(display(p.top_method))}</span>
        <span>เวลาซ่อมค่ากลาง</span><span>${medianText}</span>
        <span>Ticket ID</span><span>${escapeHtml(display(p.ticket_ids))}</span>
        <span>รายละเอียดล่าสุด</span><span>${escapeHtml(display(p.latest_detail))}</span>
      </div></div>
    </div>`)
    .addTo(map);
}

function updateMapData(records, poles) {
  if (!mapReady) return;
  map.getSource("events").setData(eventGeoJSON(records));
  map.getSource("poles").setData(poleGeoJSON(poles));
  map.getSource("districts").setData(districtGeoJSON(poles));
  rebuildDistrictMarkers(poles);
  configureMode(currentMode);
}

function fitToData(records = filteredRecords) {
  const valid = records.filter(r => r.valid_coord);
  if (!mapReady || !valid.length) return;
  const bounds = new maplibregl.LngLatBounds();
  valid.forEach(r => bounds.extend([r.lon, r.lat]));
  if (valid.length === 1) map.flyTo({center:[valid[0].lon,valid[0].lat], zoom:16});
  else map.fitBounds(bounds, {padding: 70, maxZoom: 16, duration: 700});
}

function formatIncidentDate(incident) {
  if (incident.duration_text) return incident.duration_text;
  return [incident.start_iso, incident.end_iso].filter(Boolean).join(" – ") || UNKNOWN;
}

function renderPoleDetail(key, {scrollToTop=false} = {}) {
  selectedPoleKey = key;
  const p = poleIndex.get(key);
  const container = el("poleDetail");
  if (scrollToTop) container.closest(".analytics")?.scrollTo({top:0, behavior:"smooth"});
  if (!p) {
    container.className = "pole-detail-empty";
    container.innerHTML = `เสาที่เลือกไม่อยู่ในตัวกรองปัจจุบัน`;
    return;
  }
  container.className = "";
  const history = p.incidents.map((incident, index) => `
    <div class="history-item">
      <div class="history-item-head"><strong>ครั้งที่ ${fmt.format(p.incidents.length-index)}</strong><span>${incident.is_outage?"ไฟดับ":"งานซ่อม"}</span></div>
      <div class="history-row"><span>ช่วงซ่อม</span><span>${escapeHtml(formatIncidentDate(incident))}</span></div>
      <div class="history-row"><span>เสียที่</span><span>${escapeHtml(incident.damage_type)}</span></div>
      <div class="history-row"><span>อาการ/สาเหตุ</span><span>${escapeHtml([incident.symptom, incident.details].filter(v=>v&&v!==UNKNOWN).join(" · ") || UNKNOWN)}</span></div>
      <div class="history-row"><span>วิธีซ่อม</span><span>${escapeHtml(incident.repair_method)}</span></div>
      <div class="history-row"><span>ข้อร้องเรียน</span><span>${fmt.format(incident.complaint_count)} รายการ</span></div>
    </div>`).join("");
  container.innerHTML = `
    <div class="pole-detail-head"><strong>${escapeHtml(p.pole_code)}</strong><small>${escapeHtml(p.district)} · ${escapeHtml(p.lamp_type)}</small></div>
    <div class="detail-stats">
      <div class="detail-stat"><strong>${fmt.format(p.outage_count)}</strong><span>ครั้งที่ไฟดับ</span></div>
      <div class="detail-stat"><strong>${fmt.format(p.repair_count)}</strong><span>รอบการซ่อม</span></div>
      <div class="detail-stat"><strong>${fmt.format(p.complaint_count)}</strong><span>ข้อร้องเรียน</span></div>
    </div>
    <div class="detail-meta">
      <span>เสียที่</span><span>${escapeHtml(p.damage_type)}</span>
      <span>วิธีซ่อมหลัก</span><span>${escapeHtml(p.top_method)}</span>
      <span>เวลาค่ากลาง</span><span>${p.median_hours==null?"–":fmt1.format(p.median_hours)+" ชม."}</span>
    </div>
    <div class="history-title">ประวัติการซ่อม (${fmt.format(p.incidents.length)} ครั้ง)</div>
    <div class="history-list">${history || '<div class="empty">ไม่มีประวัติการซ่อม</div>'}</div>`;
}

function focusPole(key) {
  const p = poleIndex.get(key);
  if (!p || !mapReady) return;
  if (currentMode === "heat") configureMode("outage");
  map.flyTo({center:[p.lon,p.lat], zoom:17, duration:650});
  renderPoleDetail(key, {scrollToTop:true});
}

function updateDashboard({fit=false} = {}) {
  filteredRecords = getFilteredRecords();
  const poles = aggregatePoles(filteredRecords);
  poleIndex = new Map(poles.map(p => [p.pole_key, p]));
  if (selectedPoleKey) renderPoleDetail(selectedPoleKey);
  updateKpis(filteredRecords, poles);
  const districtRepairs = poles.reduce((counts, p) => {
    counts.set(p.district, (counts.get(p.district) || 0) + p.repair_count);
    return counts;
  }, new Map());
  renderBars("districtBars", [...districtRepairs.entries()].sort((a,b)=>b[1]-a[1]), "", name => {
    setSelected("districtFilter", name); updateDashboard({fit:true});
  });
  const repairMethods = poles.flatMap(p => p.incidents).reduce((counts, incident) => {
    const method = display(incident.repair_method);
    counts.set(method, (counts.get(method) || 0) + 1);
    return counts;
  }, new Map());
  renderBars("methodBars", [...repairMethods.entries()].sort((a,b)=>b[1]-a[1]), "method", name => {
    setSelected("methodFilter", name); updateDashboard({fit:true});
  });
  updateMapData(filteredRecords, poles);
  syncRepairFilterButton();
  if (fit) fitToData(filteredRecords);
}

function resetFilters() {
  el("searchInput").value = "";
  ["districtFilter","lampFilter","damageFilter","symptomFilter","methodFilter","statusFilter"].forEach(id => el(id).value = ALL_VALUE);
  el("dateFrom").value = "__MIN_DATE__";
  el("dateTo").value = "__MAX_DATE__";
  updateDashboard({fit:true});
}

function csvCell(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"','""')}"`;
}

function exportCsv() {
  const columns = [
    ["รหัสเสาไฟฟ้า","pole_code"], ["ชนิดโคม","lamp_type"], ["lat","lat"], ["lon","lon"],
    ["ประเภทความเสียหาย","damage_type"], ["อาการที่ตรวจสอบ","symptom"], ["วิธีแก้ไข","repair_method"],
    ["รายละเอียดเพิ่มเติม","details"], ["เขต","district"], ["ticket_id","ticket_id"],
    ["สถานะข้อร้องเรียน","complaint_status"], ["ระยะเวลาดำเนินการ","duration_text"],
    ["เวลาซ่อม_ชั่วโมง","repair_hours"], ["เป็นเหตุไฟดับ","is_outage"], ["ซ่อมสำเร็จ","repair_success"]
  ];
  const lines = [columns.map(c => csvCell(c[0])).join(",")];
  filteredRecords.forEach(r => lines.push(columns.map(c => csvCell(r[c[1]])).join(",")));
  const blob = new Blob(["\ufeff" + lines.join("\r\n")], {type:"text/csv;charset=utf-8"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href=url; a.download="lighting_filtered_data.csv"; a.click();
  URL.revokeObjectURL(url);
}

// Initialize filters
fillSelect("districtFilter", "district");
fillSelect("lampFilter", "lamp_type");
fillSelect("damageFilter", "damage_type");
fillSelect("symptomFilter", "symptom");
fillSelect("methodFilter", "repair_method");
fillSelect("statusFilter", "complaint_status");
fillRepairMethodDrawer();
updateDashboard();

// MapLibre map
const map = new maplibregl.Map({
  container: "map",
  style: "https://tiles.openfreemap.org/styles/liberty",
  center: INITIAL_CENTER,
  zoom: 11,
  attributionControl: true
});
map.on("load", () => {
  mapReady = true;
  map.addSource("events", {type:"geojson", data:eventGeoJSON(filteredRecords)});
  const initialPoles = aggregatePoles(filteredRecords);
  map.addSource("poles", {type:"geojson", data:poleGeoJSON(initialPoles)});
  map.addSource("districts", {type:"geojson", data:districtGeoJSON(initialPoles)});

  map.addLayer({
    id:"repair-heat", type:"heatmap", source:"poles", maxzoom:17,
    filter:[">",["get","repair_count"],1],
    paint:{
      "heatmap-weight": ["interpolate",["linear"],["get","repair_count"],1,0.2,3,0.55,8,1.0],
      "heatmap-intensity": ["interpolate",["linear"],["zoom"],8,0.8,15,2.5],
      "heatmap-radius": ["interpolate",["linear"],["zoom"],8,16,12,27,16,44],
      "heatmap-opacity": ["interpolate",["linear"],["zoom"],8,0.85,16,0.45],
      "heatmap-color": ["interpolate",["linear"],["heatmap-density"],
        0,"rgba(43,131,186,0)",0.2,"#2b83ba",0.4,"#abdda4",0.6,"#ffffbf",0.8,"#fdae61",1,"#d7191c"]
    }
  });
  map.addLayer({
    id:"district-points", type:"circle", source:"districts", maxzoom:POLE_DETAIL_ZOOM,
    layout:{visibility:"none"},
    paint:{"circle-radius":26,"circle-color":"#155eef","circle-opacity":0.82,"circle-stroke-color":"#ffffff","circle-stroke-width":2}
  });
  map.addLayer({
    id:"district-labels", type:"symbol", source:"districts", maxzoom:POLE_DETAIL_ZOOM,
    layout:{visibility:"none","text-field":["get","district"],"text-size":11,"text-max-width":30,"text-line-height":1.15,"text-allow-overlap":false,"text-ignore-placement":false},
    paint:{"text-color":"#ffffff","text-halo-color":"rgba(16,42,67,.75)","text-halo-width":1}
  });
  map.addLayer({
    id:"pole-locations", type:"circle", source:"poles", layout:{visibility:"none"},
    paint:{"circle-radius":["interpolate",["linear"],["zoom"],9,1.5,15,4],"circle-color":"#344054","circle-opacity":0.58,"circle-stroke-color":"#ffffff","circle-stroke-width":0.6}
  });
  map.addLayer({
    id:"event-points", type:"circle", source:"events", minzoom:13,
    paint:{
      "circle-radius":["interpolate",["linear"],["zoom"],13,3,18,8],
      "circle-color":["case",["==",["get","is_outage"],1],"#d92d20","#155eef"],
      "circle-opacity":["interpolate",["linear"],["zoom"],13,0.30,17,0.85],
      "circle-stroke-color":"#ffffff","circle-stroke-width":1
    }
  });
  map.addLayer({
    id:"pole-points", type:"circle", source:"poles", minzoom:POLE_DETAIL_ZOOM, layout:{visibility:"none"},
    paint:{"circle-radius":10,"circle-color":"#d92d20","circle-opacity":0.86,"circle-stroke-color":"#ffffff","circle-stroke-width":2}
  });
  map.addLayer({
    id:"pole-labels", type:"symbol", source:"poles", minzoom:POLE_DETAIL_ZOOM, layout:{visibility:"none","text-field":["to-string",["get","outage_count"]],"text-size":11,"text-allow-overlap":true},
    paint:{"text-color":"#ffffff","text-halo-color":"rgba(0,0,0,.25)","text-halo-width":1}
  });

  map.on("click", "event-points", e => {
    const f=e.features?.[0]; if(!f) return;
    renderPoleDetail(f.properties.pole_key, {scrollToTop:true});
  });
  const openPolePopup = e => {
    const f=e.features?.[0]; if(!f) return;
    showPolePopup(f, e.lngLat);
  };
  map.on("click", "pole-points", openPolePopup);
  map.on("click", "pole-labels", openPolePopup);
  map.on("click", "pole-locations", openPolePopup);
  map.on("click", "district-points", e => {
    if (!e.features?.[0]) return;
    map.easeTo({center:e.lngLat, zoom:POLE_DETAIL_ZOOM + 0.5, duration:650});
  });
  map.on("click", "district-labels", e => {
    if (!e.features?.[0]) return;
    map.easeTo({center:e.lngLat, zoom:POLE_DETAIL_ZOOM + 0.5, duration:650});
  });
  ["event-points","pole-points","pole-labels","pole-locations","district-points","district-labels"].forEach(layer => {
    map.on("mouseenter",layer,()=>map.getCanvas().style.cursor="pointer");
    map.on("mouseleave",layer,()=>map.getCanvas().style.cursor="");
  });

  updateDashboard();
  fitToData(filteredRecords);
});
map.on("error", e => { el("mapStatus").textContent = "โหลด Basemap ไม่สำเร็จ — ตรวจอินเทอร์เน็ต"; console.error(e.error); });
map.on("zoom", updateDistrictMarkerVisibility);

// UI events
el("filterToggleButton").addEventListener("click", () => {
  const workspace = document.querySelector(".workspace");
  const collapsed = workspace.classList.toggle("filters-collapsed");
  el("filterToggleButton").textContent = collapsed ? "เปิดตัวกรอง" : "ปิดตัวกรอง";
  el("filterToggleButton").setAttribute("aria-expanded", String(!collapsed));
  requestAnimationFrame(() => map.resize());
});
el("applyButton").addEventListener("click", () => updateDashboard({fit:true}));
el("resetButton").addEventListener("click", resetFilters);
el("fitButton").addEventListener("click", () => fitToData(filteredRecords));
el("exportButton").addEventListener("click", exportCsv);
document.querySelectorAll(".mode-button").forEach(btn => btn.addEventListener("click", () => configureMode(btn.dataset.mode)));
el("heatmapToggle").addEventListener("change", refreshLayerToggles);
el("poleLocationToggle").addEventListener("change", refreshLayerToggles);
el("repairFilterOpen").addEventListener("click", () => {
  el("repairMethodDrawerSelect").value = selected("methodFilter");
  setRepairDrawerOpen(true);
});
el("repairFilterClose").addEventListener("click", () => setRepairDrawerOpen(false));
el("repairFilterBackdrop").addEventListener("click", () => setRepairDrawerOpen(false));
el("repairFilterApply").addEventListener("click", () => {
  setSelected("methodFilter", el("repairMethodDrawerSelect").value);
  updateDashboard({fit:true});
  setRepairDrawerOpen(false);
});
el("repairFilterClear").addEventListener("click", () => {
  el("repairMethodDrawerSelect").value = ALL_VALUE;
  setSelected("methodFilter", ALL_VALUE);
  updateDashboard({fit:true});
  setRepairDrawerOpen(false);
});
document.addEventListener("keydown", event => {
  if (event.key === "Escape") setRepairDrawerOpen(false);
});
["districtFilter","lampFilter","damageFilter","symptomFilter","methodFilter","statusFilter","dateFrom","dateTo"]
  .forEach(id => el(id).addEventListener("change", () => updateDashboard({fit:false})));
let searchTimer;
el("searchInput").addEventListener("input", () => { clearTimeout(searchTimer); searchTimer=setTimeout(()=>updateDashboard(),200); });
updateLegend("heat");
</script>
</body>
</html>
'''


def build_html(records: list[dict[str, Any]], quality: dict[str, Any], summary: dict[str, Any], sheet_name: str) -> str:
    def safe_json(value: Any) -> str:
        return json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

    replacements = {
        "__RECORDS_JSON__": safe_json(records),
        "__CENTER_JSON__": safe_json(summary["center"]),
        "__BOUNDS_JSON__": safe_json(summary["bounds"]),
        "__MIN_DATE__": escape(summary.get("min_date", ""), quote=True),
        "__MAX_DATE__": escape(summary.get("max_date", ""), quote=True),
    }
    html = HTML_TEMPLATE
    for key, value in replacements.items():
        html = html.replace(key, value)
    return html


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="สร้าง HTML Dashboard งานซ่อมเสาไฟด้วย MapLibre")
    parser.add_argument("--excel", required=True, help="พาธไฟล์ Excel .xlsx")
    parser.add_argument("--output", default="lighting_dashboard.html", help="ชื่อไฟล์ HTML ผลลัพธ์")
    parser.add_argument("--sheet", default=None, help="ชื่อ Worksheet (ค่าเริ่มต้นใช้ชีตแรก)")
    args = parser.parse_args()

    excel_path = Path(args.excel).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    try:
        rows, sheet_name = read_xlsx_rows(excel_path, args.sheet)
        records, quality = load_records(rows)
        summary = summarize_for_build(records)
        html = build_html(records, quality, summary, sheet_name)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(html, encoding="utf-8")
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"สร้าง Dashboard สำเร็จ: {output_path}")
    print(f"ข้อมูล: {len(records):,} แถว | พิกัดใช้ได้: {quality['valid_coordinate_rows']:,}")
    print("หลักการนับ: ยึดรหัสเสาไฟและรวมข้อร้องเรียนที่อยู่ในรอบซ่อมเดียวกัน")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
