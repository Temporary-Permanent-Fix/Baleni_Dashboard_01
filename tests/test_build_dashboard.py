from __future__ import annotations

import importlib.util
import unittest
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


PROJECT_DIR = Path(__file__).resolve().parents[1]
MODULE_PATH = PROJECT_DIR / "scripts" / "build_dashboard.py"

spec = importlib.util.spec_from_file_location("build_dashboard", MODULE_PATH)
if spec is None or spec.loader is None:  # pragma: no cover - import guard
    raise RuntimeError(f"Could not load module from {MODULE_PATH}")
build_dashboard = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build_dashboard)


class BuildDashboardTests(unittest.TestCase):
    def setUp(self) -> None:
        self._original_get_prague_now = build_dashboard.get_prague_now

    def tearDown(self) -> None:
        build_dashboard.get_prague_now = self._original_get_prague_now

    def test_freshness_marks_old_workbook_stale(self) -> None:
        build_dashboard.get_prague_now = lambda: datetime(2026, 8, 3, 8, 20, tzinfo=ZoneInfo("Europe/Prague"))
        records = [{"date": "2026-07-30"}]

        freshness = build_dashboard.build_freshness_state(records)

        self.assertEqual(freshness["expected_day"], "2026-08-02")
        self.assertEqual(freshness["latest_available_day"], "2026-07-30")
        self.assertTrue(freshness["is_stale"])

    def test_daily_kpi_uses_expected_day_not_latest_day(self) -> None:
        build_dashboard.get_prague_now = lambda: datetime(2026, 8, 3, 8, 20, tzinfo=ZoneInfo("Europe/Prague"))
        records = [
            {
                "date": "2026-07-30",
                "sheet": "SKLC3",
                "geosize": "SPO",
                "doprava": "Alzabox",
                "packing_group": "X",
                "total_count": 999,
                "ab_eliminated": 777,
            },
            {
                "date": "2026-08-02",
                "sheet": "SKLC3",
                "geosize": "SPO",
                "doprava": "Alzabox",
                "packing_group": "X",
                "total_count": 10,
                "ab_eliminated": 8,
            },
            {
                "date": "2026-07-30",
                "sheet": "CZLC4",
                "geosize": "SPO",
                "doprava": "Alzabox",
                "packing_group": "X",
                "total_count": 999,
                "ab_eliminated": 777,
            },
            {
                "date": "2026-08-02",
                "sheet": "CZLC4",
                "geosize": "SPO",
                "doprava": "Alzabox",
                "packing_group": "X",
                "total_count": 20,
                "ab_eliminated": 15,
            },
        ]
        freshness = build_dashboard.build_freshness_state(records)

        summary = build_dashboard.build_daily_kpi_summary(records, freshness=freshness)

        self.assertEqual(summary["target_day"], "2026-08-02")
        self.assertEqual(summary["sheet_rows"][0]["total_count"], 10)
        self.assertEqual(summary["sheet_rows"][0]["eliminated_count"], 8)
        self.assertEqual(summary["sheet_rows"][1]["total_count"], 20)
        self.assertEqual(summary["sheet_rows"][1]["eliminated_count"], 15)

    def test_stale_workbook_is_rejected(self) -> None:
        build_dashboard.get_prague_now = lambda: datetime(2026, 8, 3, 8, 20, tzinfo=ZoneInfo("Europe/Prague"))
        records = [{"date": "2026-07-30"}]

        with self.assertRaises(RuntimeError):
            build_dashboard.ensure_fresh_daily_data(records)


if __name__ == "__main__":
    unittest.main()
