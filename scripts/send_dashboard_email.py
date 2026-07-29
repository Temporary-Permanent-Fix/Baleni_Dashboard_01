from __future__ import annotations

import argparse
import json
import os
import re
import smtplib
import ssl
import sys
import unicodedata
from collections import defaultdict
from email.message import EmailMessage
from html import escape
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Any


PROJECT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_HTML_PATH = PROJECT_DIR / "output" / "packaging_dashboard.html"
DEFAULT_LINK = "https://balenidashboard01-xxfuafu7szdbmrxnw9wb53.streamlit.app/"
DEFAULT_TO = "peter.kadlec@alza.sk"
SPECIAL_GROUP = "spo pob dist nebalit standard"


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value).strip().lower())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def format_int(value: Any) -> str:
    return f"{int(round(float(value or 0))):,}".replace(",", " ")


def format_pct(value: Any) -> str:
    return f"{(float(value or 0) * 100):.1f}".replace(".", ",") + " %"


def is_special_elimination_group(row: dict[str, Any]) -> bool:
    return normalize_text(row.get("packing_group")) == SPECIAL_GROUP


def elimination_count(row: dict[str, Any]) -> float:
    ab = float(row.get("ab_eliminated") or 0)
    special_group = float(row.get("total_count") or 0) if is_special_elimination_group(row) else 0
    return ab + special_group


def matches_normalized(row: dict[str, Any], field: str, expected: str) -> bool:
    return normalize_text(row.get(field)) == normalize_text(expected)


def parse_record_date(value: Any) -> date | None:
    text = str(value or "").strip()
    if not text or text.lower() in {"nezadane", "none", "nan"}:
        return None
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        pass
    try:
        return datetime.strptime(text, "%d.%m.%Y").date()
    except ValueError:
        pass
    try:
        return datetime.strptime(text, "%d.%m.%y").date()
    except ValueError:
        return None


def load_payload(html_path: Path) -> dict[str, Any]:
    html = html_path.read_text(encoding="utf-8")
    match = re.search(
        r'<script id="dashboard-data" type="application/json">(?P<json>.*?)</script>',
        html,
        flags=re.S,
    )
    if not match:
        raise ValueError(f"Neviem najst dashboard data v subore: {html_path}")

    payload_text = match.group("json").replace("<\\/", "</")
    return json.loads(payload_text)


def build_summary(payload: dict[str, Any]) -> dict[str, Any]:
    records = list(payload.get("records") or [])
    target_day = date.today() - timedelta(days=1)
    daily_records = [row for row in records if parse_record_date(row.get("date")) == target_day]
    sheet_rows = []
    for sheet_name in ("SKLC3", "CZLC4"):
        filtered = [
            row
            for row in daily_records
            if matches_normalized(row, "sheet", sheet_name)
            and matches_normalized(row, "geosize", "SPO")
            and matches_normalized(row, "doprava", "Alzabox")
        ]
        total_count = sum(float(row.get("total_count") or 0) for row in filtered)
        eliminated_count = sum(elimination_count(row) for row in filtered)
        sheet_rows.append(
            {
                "sheet": sheet_name,
                "total_count": total_count,
                "eliminated_count": eliminated_count,
                "ratio": (eliminated_count / total_count) if total_count else None,
                "rows_count": len(filtered),
            }
        )

    return {
        "target_day": target_day.isoformat(),
        "sheet_rows": sheet_rows,
    }


def env_bool(name: str, default: bool = False) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if not raw:
        return default
    return raw in {"1", "true", "yes", "y", "on"}


def build_message(summary: dict[str, Any], link: str, recipient: str, sender: str) -> EmailMessage:
    subject_prefix = os.environ.get("DASHBOARD_MAIL_SUBJECT_PREFIX", "Balenie dashboard").strip() or "Balenie dashboard"
    subject = f"{subject_prefix} - {summary['target_day']}"

    lines = []
    for row in summary["sheet_rows"]:
        ratio_text = "bez dát" if row["ratio"] is None else format_pct(row["ratio"])
        lines.append(f"{row['sheet']}: {ratio_text} eliminace z geosize = SPO, doprava = alzabox")
    body_line = "\n".join(lines)
    html_body = f"""<!doctype html>
<html lang="sk">
  <body style="font-family: Arial, sans-serif; color: #15202b;">
    <p style="margin: 0; white-space: pre-line;">{escape(body_line)}</p>
  </body>
</html>"""

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = recipient
    message.set_content(body_line)
    message.add_alternative(html_body, subtype="html")
    return message


def send_email(message: EmailMessage) -> None:
    host = os.environ.get("DASHBOARD_MAIL_SMTP_HOST", "").strip()
    if not host:
        print("Mail not configured: DASHBOARD_MAIL_SMTP_HOST is empty. Skipping notification.")
        return

    port = int(os.environ.get("DASHBOARD_MAIL_SMTP_PORT", "587"))
    username = os.environ.get("DASHBOARD_MAIL_SMTP_USERNAME", "").strip()
    password = os.environ.get("DASHBOARD_MAIL_SMTP_PASSWORD", "")
    use_ssl = env_bool("DASHBOARD_MAIL_USE_SSL", False)
    use_starttls = env_bool("DASHBOARD_MAIL_USE_STARTTLS", True)

    context = ssl.create_default_context()
    if use_ssl:
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as smtp:
            if username:
                smtp.login(username, password)
            smtp.send_message(message)
        return

    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.ehlo()
        if use_starttls:
            smtp.starttls(context=context)
            smtp.ehlo()
        if username:
            smtp.login(username, password)
        smtp.send_message(message)


def main() -> int:
    parser = argparse.ArgumentParser(description="Send a dashboard update email.")
    parser.add_argument("--html", dest="html_path", default=str(DEFAULT_HTML_PATH))
    parser.add_argument("--to", dest="recipient", default=os.environ.get("DASHBOARD_MAIL_TO", DEFAULT_TO))
    parser.add_argument("--from", dest="sender", default=os.environ.get("DASHBOARD_MAIL_FROM", ""))
    args = parser.parse_args()

    html_path = Path(args.html_path)
    if not html_path.is_absolute():
        html_path = (PROJECT_DIR / html_path).resolve()

    if not html_path.exists():
        print(f"Dashboard HTML not found: {html_path}")
        return 1

    summary = build_summary(load_payload(html_path))
    sender = args.sender.strip() or os.environ.get("DASHBOARD_MAIL_SMTP_USERNAME", "").strip()
    if not sender:
        print("Mail not configured: DASHBOARD_MAIL_FROM or DASHBOARD_MAIL_SMTP_USERNAME is empty. Skipping notification.")
        return 0

    message = build_message(summary, DEFAULT_LINK, args.recipient.strip(), sender)
    send_email(message)
    print(f"Mail sent to {args.recipient.strip()} with daily KPI for {summary['target_day']}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover - surfaced in refresh logs
        print(f"Mail notification failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
