from __future__ import annotations

import json
import math
import re
import subprocess
import sys
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "packaging_dashboard.html"

TARGET_RATIO = 0.75

COLUMN_ALIASES = {
    "date": ["datum", "date", "day", "den"],
    "geosize": ["geosize", "geo size", "velikost", "size"],
    "station": ["stanice baleni", "stanica balenia", "packing station", "station"],
    "packing_group": ["baliaca skupina", "balici skupina", "packing group", "packaging group", "skupina"],
    "doprava": ["doprava", "detail dopravy", "transport detail", "delivery detail"],
    "ab_eliminated": ["ab eliminovane", "ab eliminated", "eliminovane", "eliminated"],
    "piece_count": ["pocet kusu", "počet kusů", "pieces", "quantity", "mnozstvo", "mnozstvi", "množstvo", "množství", "kusy"],
    "total_count": ["celkovy pocet", "total count", "count", "pocet", "pocet jobline", "storejoblines", "store job lines"],
}

GEOSIZE_VALUES = {"SPO", "BPO", "XPO", "XL", "VB"}
STATION_VALUES = {"EXPRESS", "EXPRES", "L40", "MO", "SO01", "SOA1", "SOA0", "XXL"}
JOBLINE_MEASURES = {"pocet jobline", "jobline", "storejoblines", "store job lines"}
PIECE_MEASURES = {"pocet kusu", "počet kusů", "pieces", "quantity", "mnozstvo", "mnozstvi", "množstvo", "množství", "kusy"}


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value).strip().lower())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def find_excel_file() -> Path:
    excel_files = [
        path
        for path in INPUT_DIR.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls", ".xlsm"}
        and not path.name.startswith("~$")
    ]
    if not excel_files:
        raise FileNotFoundError(
            "V zlozke input nie je ziaden Excel subor. Vloz tam .xlsx, .xls alebo .xlsm subor."
        )
    return sorted(excel_files, key=lambda path: path.stat().st_mtime, reverse=True)[0]


def copy_excel_for_reading(excel_path: Path) -> Path:
    temp_dir = Path(tempfile.gettempdir())
    copy_path = temp_dir / f"dashboard-{excel_path.stem}-{excel_path.stat().st_mtime_ns}.xlsx"
    result = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            f"Copy-Item -LiteralPath '{str(excel_path)}' -Destination '{str(copy_path)}' -Force",
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise PermissionError(result.stderr.strip() or result.stdout.strip())
    return copy_path


def infer_columns(columns: list[Any]) -> dict[str, str | None]:
    normalized_columns = {column: normalize_text(column) for column in columns}
    result: dict[str, str | None] = {}
    for field, aliases in COLUMN_ALIASES.items():
        normalized_aliases = [normalize_text(alias) for alias in aliases]
        exact = next(
            (
                column
                for column, normalized in normalized_columns.items()
                if normalized in normalized_aliases
            ),
            None,
        )
        if exact is not None:
            result[field] = str(exact)
            continue
        partial = next(
            (
                column
                for column, normalized in normalized_columns.items()
                if any(alias in normalized or normalized in alias for alias in normalized_aliases)
            ),
            None,
        )
        result[field] = str(partial) if partial is not None else None
    return result


def safe_number(value: Any, default: float = 0) -> float:
    if pd.isna(value):
        return default
    if isinstance(value, (int, float)):
        return 0 if math.isnan(value) else float(value)
    text = str(value).strip().lower().replace(",", ".")
    if not text:
        return default
    try:
        return float(text)
    except ValueError:
        return default


def as_dimension(value: Any, fallback: str = "Nezadane") -> str:
    if pd.isna(value):
        return fallback
    text = str(value).strip()
    return text if text else fallback


def as_date(value: Any) -> str:
    if pd.isna(value):
        return "Nezadane"
    parsed = parse_date_value(value)
    if parsed is None:
        return as_dimension(value)
    return parsed


def parse_date_value(value: Any) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    dayfirst = not bool(re.match(r"^\d{4}-\d{1,2}-\d{1,2}$", text))
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=dayfirst)
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


def is_numeric_value(value: Any) -> bool:
    if pd.isna(value):
        return False
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def get_cell_indent(cell: Any) -> int:
    indent = getattr(getattr(cell, "alignment", None), "indent", 0) or 0
    try:
        return int(indent)
    except (TypeError, ValueError):
        return 0


def find_pivot_header_row(worksheet: Any) -> int | None:
    for row_index in range(1, worksheet.max_row + 1):
        value = worksheet.cell(row_index, 1).value
        if "popisky radku" in normalize_text(value):
            return row_index
    return None


def measure_kind(value: Any) -> str | None:
    normalized = normalize_text(value)
    if normalized in {normalize_text(item) for item in JOBLINE_MEASURES}:
        return "jobline"
    if normalized in {normalize_text(item) for item in PIECE_MEASURES}:
        return "piece"
    return None


def find_pivot_date_row(worksheet: Any, header_row: int) -> int | None:
    best_row: int | None = None
    best_score = 0
    for row_index in range(1, header_row + 1):
        score = 0
        for column_index in range(2, worksheet.max_column + 1):
            if parse_date_value(worksheet.cell(row_index, column_index).value) is not None:
                score += 1
        if score > best_score:
            best_row = row_index
            best_score = score
    return best_row if best_score >= 2 else None


def find_pivot_measure_row(worksheet: Any, date_row: int, header_row: int) -> int | None:
    upper_bound = min(worksheet.max_row, header_row + 1)
    best_row: int | None = None
    best_score = 0
    for row_index in range(date_row + 1, upper_bound + 1):
        score = 0
        non_empty = 0
        for column_index in range(2, worksheet.max_column + 1):
            value = worksheet.cell(row_index, column_index).value
            if pd.isna(value) or value is None or str(value).strip() == "":
                continue
            non_empty += 1
            if measure_kind(value) is not None:
                score += 1
        if score > best_score and score >= 2 and score >= max(2, non_empty // 2):
            best_row = row_index
            best_score = score
    return best_row


def parse_pivot_sheet(worksheet: Any, sheet_name: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    header_row = find_pivot_header_row(worksheet)
    if header_row is None:
        return [], {}

    date_row = find_pivot_date_row(worksheet, header_row)
    if date_row is None:
        return [], {}
    measure_row = find_pivot_measure_row(worksheet, date_row, header_row)

    date_by_column: dict[int, str] = {}
    current_date: str | None = None
    for column_index in range(2, worksheet.max_column + 1):
        parsed = parse_date_value(worksheet.cell(date_row, column_index).value)
        if parsed is not None:
            current_date = parsed
        if current_date is not None:
            date_by_column[column_index] = current_date

    measure_by_column: dict[int, str | None] = {}
    if measure_row is not None:
        for column_index in range(2, worksheet.max_column + 1):
            measure_by_column[column_index] = measure_kind(worksheet.cell(measure_row, column_index).value)

    records_map: dict[tuple[str, str, str, str, str, str], dict[str, Any]] = {}
    current_doprava = "Nezadane"
    current_geosize = "Nezadane"
    current_station = "Nezadane"
    current_packing_group = "Nezadane"

    start_row = (measure_row or date_row) + 1

    for row_index in range(start_row, worksheet.max_row + 1):
        label = as_dimension(worksheet.cell(row_index, 1).value)
        if label == "Nezadane":
            continue
        normalized_label = normalize_text(label)
        if normalized_label in {"celkovy soucet", "grand total"}:
            continue

        indent = get_cell_indent(worksheet.cell(row_index, 1))
        row_has_values = any(
            is_numeric_value(worksheet.cell(row_index, column_index).value)
            for column_index in range(2, worksheet.max_column + 1)
        )
        upper_label = label.upper()

        if indent == 0 and not row_has_values and upper_label not in GEOSIZE_VALUES and upper_label not in STATION_VALUES:
            current_doprava = label
            current_geosize = "Nezadane"
            current_station = "Nezadane"
            current_packing_group = "Nezadane"
            continue
        if upper_label in GEOSIZE_VALUES:
            current_geosize = upper_label
            current_station = "Nezadane"
            current_packing_group = "Nezadane"
            continue
        if upper_label in STATION_VALUES:
            current_station = "Expres" if upper_label in {"EXPRESS", "EXPRES"} else upper_label
            current_packing_group = "Nezadane"
            continue
        if indent >= 3 and not row_has_values:
            current_packing_group = label
            continue
        if not row_has_values:
            continue

        doprava = current_doprava
        row_group = current_packing_group if indent >= 4 else label

        for column_index in range(2, worksheet.max_column + 1):
            value = safe_number(worksheet.cell(row_index, column_index).value)
            if value <= 0:
                continue
            date_value = date_by_column.get(column_index)
            if date_value is None:
                continue
            kind = measure_by_column.get(column_index) if measure_row is not None else "jobline"
            key = (date_value, current_geosize, current_station, row_group, doprava, sheet_name)
            record = records_map.setdefault(
                key,
                {
                    "date": date_value,
                    "geosize": current_geosize,
                    "station": current_station,
                    "packing_group": row_group,
                    "doprava": doprava,
                    "ab_eliminated": 0.0,
                    "total_count": 0.0,
                    "piece_count": 0.0,
                    "sheet": sheet_name,
                },
            )
            if kind == "piece":
                record["piece_count"] += value
            else:
                record["total_count"] += value
                if "ab eliminovane" in normalized_label:
                    record["ab_eliminated"] += value

    records = list(records_map.values())

    detected = {
        "layout": "pivot",
        "date": "Popisky sloupcu",
        "geosize": "Popisky radku uroven 1",
        "station": "Popisky radku uroven 2",
        "packing_group": "Popisky radku uroven 3",
        "doprava": "Doprava",
        "ab_eliminated": "AB Eliminovane etikety",
        "total_count": "Pocet jobline",
        "piece_count": "Pocet kusu",
    }
    return records, detected


def load_records(excel_path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    sheets = pd.read_excel(excel_path, sheet_name=None)
    workbook = load_workbook(excel_path, data_only=True)
    records: list[dict[str, Any]] = []
    sheet_info: list[dict[str, Any]] = []
    detected_columns: dict[str, Any] = {}

    for sheet_name, frame in sheets.items():
        if frame.empty:
            continue

        inferred = infer_columns(list(frame.columns))
        standard_layout = bool(inferred.get("ab_eliminated") or inferred.get("total_count") or inferred.get("piece_count"))
        if not standard_layout:
            pivot_records, pivot_detected = parse_pivot_sheet(workbook[sheet_name], sheet_name)
            if pivot_records:
                records.extend(pivot_records)
                raw_sheet = workbook[sheet_name]
                sheet_info.append(
                    {
                        "sheet": sheet_name,
                        "rows": raw_sheet.max_row,
                        "columns": raw_sheet.max_column,
                        "detected": pivot_detected,
                    }
                )
                detected_columns = detected_columns or pivot_detected
                continue

        sheet_info.append(
            {
                "sheet": sheet_name,
                "rows": len(frame),
                "columns": len(frame.columns),
                "detected": inferred,
            }
        )
        detected_columns = detected_columns or inferred

        for _, row in frame.iterrows():
            total_column = inferred.get("total_count")
            eliminated_column = inferred.get("ab_eliminated")
            piece_column = inferred.get("piece_count")
            total = safe_number(row.get(total_column), 1) if total_column else 1
            eliminated = safe_number(row.get(eliminated_column), 0) if eliminated_column else 0
            pieces = safe_number(row.get(piece_column), 0) if piece_column else 0
            if total <= 0 and eliminated <= 0:
                continue
            records.append(
                {
                    "date": as_date(row.get(inferred["date"])) if inferred.get("date") else "Nezadane",
                    "geosize": as_dimension(row.get(inferred["geosize"])) if inferred.get("geosize") else "Nezadane",
                    "station": as_dimension(row.get(inferred["station"])) if inferred.get("station") else "Nezadane",
                    "packing_group": as_dimension(row.get(inferred["packing_group"])) if inferred.get("packing_group") else "Nezadane",
                    "doprava": as_dimension(row.get(inferred["doprava"])) if inferred.get("doprava") else "Nezadane",
                    "ab_eliminated": eliminated,
                    "total_count": total,
                    "piece_count": pieces,
                    "sheet": sheet_name,
                }
            )

    metadata = {
        "source_file": excel_path.name,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "target_ratio": TARGET_RATIO,
        "sheet_info": sheet_info,
        "detected_columns": detected_columns,
    }
    return records, metadata


def build_payload(records: list[dict[str, Any]], metadata: dict[str, Any]) -> dict[str, Any]:
    return {"metadata": metadata, "records": records}


def to_json_for_html(payload: dict[str, Any]) -> str:
    return json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")


def render_html(payload: dict[str, Any]) -> str:
    data_json = to_json_for_html(payload)
    html = """<!doctype html>
<html lang="sk">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Vyvoj balenia dashboard</title>
  <style>
    :root {
      --bg: #f5f7fa;
      --panel: #ffffff;
      --ink: #15202b;
      --muted: #5f6b7a;
      --line: #d8dee8;
      --blue: #2563eb;
      --green: #0f9f6e;
      --amber: #d97706;
      --red: #dc2626;
      --shadow: 0 12px 30px rgba(23, 37, 84, .08);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: "Segoe UI", Arial, sans-serif;
    }
    header {
      background: #111827;
      color: white;
      padding: 24px clamp(16px, 3vw, 36px);
    }
    header h1 {
      margin: 0 0 8px;
      font-size: clamp(24px, 3vw, 36px);
      letter-spacing: 0;
    }
    header p { margin: 0; color: #cbd5e1; }
    main {
      width: min(1440px, 100%);
      margin: 0 auto;
      padding: 24px clamp(12px, 2.5vw, 32px) 36px;
    }
    .meta {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      color: var(--muted);
      font-size: 13px;
      margin-bottom: 16px;
    }
    .pill {
      background: #eaf0f8;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 6px 10px;
    }
    .dashboards {
      display: grid;
      gap: 22px;
    }
    .dashboard {
      display: grid;
      gap: 14px;
      padding: 18px;
      background: #eef2f7;
      border: 1px solid var(--line);
      border-radius: 8px;
    }
    .dashboard-head {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 16px;
    }
    .dashboard-head h2 {
      margin: 0 0 4px;
      font-size: 20px;
      letter-spacing: 0;
    }
    .dashboard-head p {
      margin: 0;
      color: var(--muted);
      font-size: 13px;
    }
    .filters {
      display: grid;
      grid-template-columns: repeat(5, minmax(160px, 1fr));
      gap: 12px;
    }
    label {
      display: grid;
      gap: 6px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
    }
    select {
      min-height: 38px;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 7px 10px;
      background: white;
      color: var(--ink);
      font-size: 14px;
      width: 100%;
    }
    .cards {
      display: grid;
      grid-template-columns: repeat(4, minmax(170px, 1fr));
      gap: 14px;
    }
    .card, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .card { padding: 16px; }
    .card span {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      text-transform: uppercase;
    }
    .card strong {
      display: block;
      margin-top: 8px;
      font-size: clamp(24px, 2.8vw, 34px);
    }
    .card small { color: var(--muted); }
    .goalbar {
      height: 9px;
      background: #e5e7eb;
      border-radius: 999px;
      overflow: hidden;
      margin-top: 12px;
    }
    .goalbar div {
      height: 100%;
      width: 0;
      background: var(--green);
      transition: width .2s ease;
    }
    .grid {
      display: grid;
      grid-template-columns: minmax(0, 1.35fr) minmax(340px, .65fr);
      gap: 16px;
      align-items: start;
    }
    .panel {
      padding: 16px;
      min-width: 0;
    }
    .panel h2 {
      margin: 0 0 14px;
      font-size: 18px;
      letter-spacing: 0;
    }
    .chart-wrap {
      width: 100%;
      height: 360px;
    }
    svg { width: 100%; height: 100%; display: block; }
    .table-wrap { overflow: auto; }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th, td {
      padding: 9px 8px;
      border-bottom: 1px solid var(--line);
      text-align: right;
      white-space: nowrap;
    }
    th:first-child, td:first-child { text-align: left; }
    th {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
    }
    .ratio.good { color: var(--green); font-weight: 700; }
    .ratio.mid { color: var(--amber); font-weight: 700; }
    .ratio.bad { color: var(--red); font-weight: 700; }
    .empty {
      padding: 28px;
      text-align: center;
      color: var(--muted);
      border: 1px dashed var(--line);
      border-radius: 8px;
      background: #fbfcfe;
    }
    @media (max-width: 1020px) {
      .filters, .cards, .grid { grid-template-columns: 1fr 1fr; }
      .grid .panel:first-child { grid-column: 1 / -1; }
      .dashboard-head { flex-direction: column; }
    }
    @media (max-width: 680px) {
      .filters, .cards, .grid { grid-template-columns: 1fr; }
      header { padding-top: 18px; padding-bottom: 18px; }
      .chart-wrap { height: 300px; }
    }
  </style>
</head>
<body>
  <header>
    <h1>Vyvoj balenia</h1>
    <p>AB eliminated / total count, target 75 %</p>
  </header>
  <main>
    <div class="meta" id="meta"></div>
    <div class="dashboards" id="dashboards"></div>
  </main>
  <script id="dashboard-data" type="application/json">__DATA_JSON__</script>
  <script>
    const payload = JSON.parse(document.getElementById('dashboard-data').textContent);
    const records = payload.records || [];
    const targetRatio = payload.metadata?.target_ratio ?? 0.75;
    const balikovkaTarget = 0.10;
    const sheetOrder = (payload.metadata?.sheet_info || []).map(info => info.sheet).filter(Boolean);
    const dashboardSheets = sheetOrder.length
      ? sheetOrder
      : [...new Set(records.map(row => row.sheet).filter(Boolean))];
    const balikovkaTransports = {
      SKLC3: [
        'DHL',
        'DPD',
        'B2B',
        'AlzaExpres',
        'ExpressOne',
        'FoxPost',
        'Gebruder Weiss',
        'GO!',
        'Maďarská pošta',
        'MyFlexBox',
        'Post AT',
        'Pošta',
        'Pošta Slovensko',
        'PPL',
        'SPS',
        'TopTrans',
        'WeDo',
        'Zásilkovna',
      ],
      CZLC4: [
        'DHL',
        'DPD',
        'B2B',
        'AlzaExpres',
        'ExpressOne',
        'FoxPost',
        'Gebruder Weiss',
        'GO!',
        'Kurýr',
        'Maďarská pošta',
        'MyFlexBox',
        'Post AT',
        'Pošta',
        'Pošta Slovensko',
        'PPL',
        'SPS',
        'TopTrans',
        'WeDo',
        'Zásilkovna',
      ],
    };

    const labels = {
      date: 'Date',
      geosize: 'Geosize',
      station: 'Packing station',
      packing_group: 'Packing group',
      doprava: 'Transport detail',
    };

    const fmtInt = new Intl.NumberFormat('sk-SK', { maximumFractionDigits: 0 });
    const fmtPct = new Intl.NumberFormat('sk-SK', {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    });

    function formatInt(value) {
      return fmtInt.format(Math.round(Number(value) || 0));
    }

    function formatPct(value) {
      return fmtPct.format((Number(value) || 0) * 100) + ' %';
    }

    function ratioClass(value) {
      if (value >= targetRatio) return 'good';
      if (value >= 0.65) return 'mid';
      return 'bad';
    }

    function balikovkaClass(value) {
      if (value <= balikovkaTarget) return 'good';
      if (value <= 0.12) return 'mid';
      return 'bad';
    }

    function slugify(value) {
      return String(value ?? '')
        .toLowerCase()
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '');
    }

    function normalizeText(value) {
      return String(value ?? '')
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
    }

    function escapeHtml(value) {
      return String(value).replace(/[&<>"']/g, char => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
      }[char]));
    }

    function normalizeLabel(value) {
      return String(value ?? 'Nezadane');
    }

    function displayLabel(field, value) {
      const normalized = normalizeLabel(value);
      if (field === 'doprava' && normalized === 'Nezadane') return 'Pobocka';
      return normalized;
    }

    function isSpecialEliminationGroup(row) {
      return normalizeLabel(row.packing_group).toLowerCase() === 'spo pob, dist nebalit standard';
    }

    function eliminationCount(row) {
      const ab = Number(row.ab_eliminated) || 0;
      const specialGroup = isSpecialEliminationGroup(row) ? (Number(row.total_count) || 0) : 0;
      return ab + specialGroup;
    }

    function balikovkaTransportSet(sheetName) {
      return new Set((balikovkaTransports[sheetName] || []).map(normalizeText));
    }

    function balikovkaCount(rows, sheetName) {
      const allowed = balikovkaTransportSet(sheetName);
      if (!allowed.size) return 0;
      return rows.reduce((sum, row) => {
        const transport = normalizeText(row.doprava);
        return sum + (allowed.has(transport) ? (Number(row.total_count) || 0) : 0);
      }, 0);
    }

    function pieceCount(rows) {
      return rows.reduce((sum, row) => sum + (Number(row.piece_count) || 0), 0);
    }

    function aggregate(rows) {
      let eliminated = 0;
      let total = 0;
      for (const row of rows) {
        eliminated += eliminationCount(row);
        total += Number(row.total_count) || 0;
      }
      return {
        ab: eliminated,
        total,
        ratio: total ? eliminated / total : 0,
      };
    }

    function allValues(rows, field) {
      return ['all', ...new Set(rows.map(row => normalizeLabel(row[field])))].sort((a, b) => a.localeCompare(b, 'sk'));
    }

    function seriesByDate(rows) {
      const map = new Map();
      for (const row of rows) {
        const key = normalizeLabel(row.date);
        const item = map.get(key) || { key, ab: 0, total: 0 };
        item.ab += eliminationCount(row);
        item.total += Number(row.total_count) || 0;
        map.set(key, item);
      }
      return [...map.values()].sort((a, b) => a.key.localeCompare(b.key, 'sk')).map(item => ({
        ...item,
        ratio: item.total ? item.ab / item.total : 0,
      }));
    }

    function renderTrend(rows, targetId) {
      const data = seriesByDate(rows);
      const holder = document.getElementById(targetId);
      if (!data.length) {
        holder.innerHTML = '<div class="empty">No data for trend.</div>';
        return;
      }

      const width = 920;
      const height = 340;
      const left = 54;
      const right = 22;
      const top = 18;
      const bottom = 48;
      const innerW = width - left - right;
      const innerH = height - top - bottom;
      const maxY = Math.max(targetRatio, ...data.map(d => d.ratio), 0.8);
      const x = index => left + (data.length === 1 ? innerW / 2 : index * innerW / (data.length - 1));
      const y = value => top + innerH - (value / maxY) * innerH;
      const points = data.map((d, i) => `${x(i)},${y(d.ratio)}`).join(' ');
      const targetY = y(targetRatio);
      const labelsLine = data.map((d, i) => {
        const show = data.length <= 12 || i === 0 || i === data.length - 1 || i % Math.ceil(data.length / 8) === 0;
        return show ? `<text x="${x(i)}" y="${height - 16}" text-anchor="middle" font-size="11" fill="#5f6b7a">${escapeHtml(String(d.key).slice(5))}</text>` : '';
      }).join('');

      holder.innerHTML = `
        <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Trend AB eliminated">
          <line x1="${left}" y1="${top}" x2="${left}" y2="${height - bottom}" stroke="#d8dee8"/>
          <line x1="${left}" y1="${height - bottom}" x2="${width - right}" y2="${height - bottom}" stroke="#d8dee8"/>
          <line x1="${left}" y1="${targetY}" x2="${width - right}" y2="${targetY}" stroke="#0f9f6e" stroke-dasharray="6 6"/>
          <text x="${width - right}" y="${targetY - 8}" text-anchor="end" font-size="12" fill="#0f9f6e">target 75 %</text>
          <polyline fill="none" stroke="#2563eb" stroke-width="3" points="${points}"/>
          ${data.map((d, i) => `<circle cx="${x(i)}" cy="${y(d.ratio)}" r="4" fill="#2563eb"><title>${escapeHtml(d.key)}: ${formatPct(d.ratio)}</title></circle>`).join('')}
          ${labelsLine}
        </svg>
      `;
    }

    function renderTable(id, rows, key) {
      const map = new Map();
      let totalVolume = 0;
      for (const row of rows) {
        totalVolume += Number(row.total_count) || 0;
        const name = normalizeLabel(row[key]);
        const item = map.get(name) || { name, ab: 0, total: 0 };
        item.ab += eliminationCount(row);
        item.total += Number(row.total_count) || 0;
        map.set(name, item);
      }
      const data = [...map.values()].map(item => ({
        ...item,
        share: totalVolume ? item.total / totalVolume : 0,
        elimRatio: item.total ? item.ab / item.total : 0,
      })).sort((a, b) => b.total - a.total);
      const table = document.getElementById(id);
      if (!data.length) {
        table.innerHTML = '<tbody><tr><td class="empty">No data for the selected filter.</td></tr></tbody>';
        return;
      }
      table.innerHTML = `
        <thead><tr><th>${labels[key]}</th><th>SJLs</th><th>Share</th><th>Elimination</th></tr></thead>
        <tbody>
          ${data.map(row => `
            <tr>
              <td>${escapeHtml(displayLabel(key, row.name))}</td>
              <td>${formatInt(row.total)}</td>
              <td>${formatPct(row.share)}</td>
              <td class="ratio ${ratioClass(row.elimRatio)}">${formatPct(row.elimRatio)}</td>
            </tr>
          `).join('')}
        </tbody>
      `;
    }

    function renderVolumeTable(id, rows, key) {
      renderTable(id, rows, key);
    }

    function renderDashboard(sheetName, index) {
      const info = (payload.metadata?.sheet_info || []).find(item => item.sheet === sheetName) || {};
      const sheetId = slugify(sheetName) || `sheet-${index + 1}`;
      const sheetRows = records.filter(row => normalizeLabel(row.sheet) === normalizeLabel(sheetName));
      const prefix = `${sheetId}_`;
      const root = document.createElement('section');
      root.className = 'dashboard';
      root.id = `dashboard_${sheetId}`;
      root.innerHTML = `
        <div class="dashboard-head">
          <div>
            <h2>Dashboard ${index + 1} - ${escapeHtml(sheetName)}</h2>
            <p>Same conditions as the first dashboard, rendered from the same workbook.</p>
          </div>
          <div class="meta" id="${prefix}meta"></div>
        </div>
        <section class="filters" id="${prefix}filters"></section>
        <section class="cards">
          <div class="card">
            <span>Elimination share</span>
            <strong id="${prefix}kpiRatio">0 %</strong>
            <small>Target: 75 %</small>
            <div class="goalbar"><div id="${prefix}goalFill"></div></div>
          </div>
          <div class="card">
            <span>Balikovka</span>
            <strong id="${prefix}kpiBalikovka">0 %</strong>
            <small id="${prefix}kpiBalikovkaCount">0 SJLs</small>
            <small>Target: &lt; 10 % of SJLs</small>
          </div>
          <div class="card">
            <span>Počet kusov</span>
            <strong id="${prefix}kpiPieces">0</strong>
            <small id="${prefix}kpiPiecesShare">0 % z kusov</small>
            <small>From Počet kusů values</small>
          </div>
          <div class="card"><span>Eliminated</span><strong id="${prefix}kpiEliminated">0</strong><small>StoreJobLines</small></div>
          <div class="card"><span>Selected total</span><strong id="${prefix}kpiSelectedTotal">0</strong><small>StoreJobLines</small></div>
          <div class="card"><span>Dashboard total</span><strong id="${prefix}kpiTotal">0</strong><small>StoreJobLines</small></div>
          <div class="card"><span>Gap to target</span><strong id="${prefix}kpiGap">0 b.</strong><small>percentage points</small></div>
        </section>
        <section class="grid">
          <div class="panel">
            <h2>Trend by day</h2>
            <div class="chart-wrap" id="${prefix}trendChart"></div>
          </div>
          <div class="panel">
            <h2>Geosize</h2>
            <div class="table-wrap"><table id="${prefix}geosizeTable"></table></div>
          </div>
          <div class="panel">
            <h2>Packing station</h2>
            <div class="table-wrap"><table id="${prefix}stationTable"></table></div>
          </div>
          <div class="panel">
            <h2>Packing groups</h2>
            <div class="table-wrap"><table id="${prefix}groupTable"></table></div>
          </div>
          <div class="panel">
            <h2>Transport detail</h2>
            <div class="table-wrap"><table id="${prefix}detailTable"></table></div>
          </div>
        </section>
      `;
      document.getElementById('dashboards').appendChild(root);

      const state = {
        date: 'all',
        geosize: 'all',
        station: 'all',
        packing_group: 'all',
        doprava: 'all',
      };

      function filteredRows() {
        return sheetRows.filter(row => {
          return Object.entries(state).every(([field, selected]) => {
            if (selected === 'all') return true;
            return normalizeLabel(row[field]) === selected;
          });
        });
      }

      function renderMeta(rows) {
        const meta = payload.metadata || {};
        const totalRows = sheetRows.length;
        const detectionLabel = info.detected?.piece_count
          ? 'AB + Total + Pieces + transport'
          : (info.detected?.doprava ? 'AB + Total + transport' : (info.detected?.ab_eliminated ? 'AB + Total' : 'pivot/flat'));
        document.getElementById(`${prefix}meta`).innerHTML = [
          `Source: ${meta.source_file || 'unknown'}`,
          `Sheet: ${sheetName}`,
          `Generated: ${meta.generated_at || ''}`,
          `Rows: ${formatInt(rows.length)} / ${formatInt(totalRows)}`,
          `Detected: ${detectionLabel}`,
        ].map(text => `<span class="pill">${text}</span>`).join('');
      }

      function renderFilters() {
        const holder = document.getElementById(`${prefix}filters`);
        holder.innerHTML = Object.keys(labels).map(field => `
          <label>${labels[field]}<select id="${prefix}filter_${field}"></select></label>
        `).join('');
        for (const field of Object.keys(state)) {
          const select = document.getElementById(`${prefix}filter_${field}`);
          if (!select) continue;
          select.innerHTML = allValues(sheetRows, field).map(value => {
            const label = value === 'all' ? 'All' : displayLabel(field, value);
            return `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`;
          }).join('');
          select.value = state[field];
          select.onchange = event => {
            state[field] = event.target.value;
            render();
          };
        }
      }

      function renderKpis(rows) {
        const summary = aggregate(rows);
        const dashboardSummary = aggregate(sheetRows);
        const balikovka = balikovkaCount(rows, sheetName);
        const balikovkaRatio = summary.total ? balikovka / summary.total : 0;
        const selectedPieces = pieceCount(rows);
        const dashboardPieces = pieceCount(sheetRows);
        const pieceRatio = dashboardPieces ? selectedPieces / dashboardPieces : 0;
        document.getElementById(`${prefix}kpiRatio`).textContent = formatPct(summary.ratio);
        document.getElementById(`${prefix}kpiRatio`).className = 'ratio ' + ratioClass(summary.ratio);
        document.getElementById(`${prefix}kpiBalikovka`).textContent = formatPct(balikovkaRatio);
        document.getElementById(`${prefix}kpiBalikovka`).className = 'ratio ' + balikovkaClass(balikovkaRatio);
        document.getElementById(`${prefix}kpiBalikovkaCount`).textContent = `${formatInt(balikovka)} SJLs`;
        document.getElementById(`${prefix}kpiPieces`).textContent = formatInt(selectedPieces);
        document.getElementById(`${prefix}kpiPiecesShare`).textContent = `${formatPct(pieceRatio)} z kusov`;
        document.getElementById(`${prefix}kpiEliminated`).textContent = formatInt(summary.ab);
        document.getElementById(`${prefix}kpiSelectedTotal`).textContent = formatInt(summary.total);
        document.getElementById(`${prefix}kpiTotal`).textContent = formatInt(dashboardSummary.total);
        document.getElementById(`${prefix}kpiGap`).textContent = ((summary.ratio - targetRatio) * 100).toLocaleString('sk-SK', {
          minimumFractionDigits: 1,
          maximumFractionDigits: 1,
        }) + ' b.';
        document.getElementById(`${prefix}goalFill`).style.width = Math.min(100, (summary.ratio / targetRatio) * 100) + '%';
      }

      function render() {
        const rows = filteredRows();
        renderMeta(rows);
        renderKpis(rows);
        renderFilters();
        renderTrend(rows, `${prefix}trendChart`);
        renderTable(`${prefix}geosizeTable`, rows, 'geosize');
        renderTable(`${prefix}stationTable`, rows, 'station');
        renderTable(`${prefix}groupTable`, rows, 'packing_group');
        renderVolumeTable(`${prefix}detailTable`, rows, 'doprava');
      }

      render();
    }

    document.getElementById('meta').innerHTML = [
      `Source: ${payload.metadata?.source_file || 'unknown'}`,
      `Generated: ${payload.metadata?.generated_at || ''}`,
      `Dashboards: ${dashboardSheets.length || 0}`,
    ].map(text => `<span class="pill">${text}</span>`).join('');

    const dashboardsHolder = document.getElementById('dashboards');
    if (!dashboardSheets.length) {
      dashboardsHolder.innerHTML = '<div class="empty">No records found in the input workbook.</div>';
    } else {
      dashboardSheets.forEach((sheetName, index) => renderDashboard(sheetName, index));
    }
  </script>
</body>
</html>
"""
    return html.replace("__DATA_JSON__", data_json)


def save_dashboard(payload: dict[str, Any]) -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_FILE.write_text(render_html(payload), encoding="utf-8")


def main() -> None:
    print("Startujem tvorbu dashboardu...")
    try:
        excel_path = find_excel_file()
    except FileNotFoundError as error:
        print(f"Chyba: {error}")
        print(f"Upload priecinok: {INPUT_DIR}")
        sys.exit(1)

    print(f"Nacitavam subor: {excel_path.name}")
    readable_copy = copy_excel_for_reading(excel_path)
    try:
        records, metadata = load_records(readable_copy)
    finally:
        try:
            readable_copy.unlink(missing_ok=True)
        except OSError:
            pass

    payload = build_payload(records, metadata)
    save_dashboard(payload)

    print("Hotovo.")
    print(f"Dashboard je ulozeny tu: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
