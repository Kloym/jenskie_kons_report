"""Run: python -m unittest discover -s tests -p test_service_actuals.py -v"""
import copy
import json
import sys
import tempfile
import unittest
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from zhk_actuals import MonthlyServiceActuals, snapshot_from_services
from zhk_annual_plan import AnnualPlanError, snapshot
import zhk_dashboard as dashboard


class ActualsTests(unittest.TestCase):
    def setUp(self):
        self.config = {"version": 1, "asOf": "2026-10-22", "monthPlanTotal": 99,
                       "departments": [{"name": "ЖК №10", "yearPlan": 1000,
                                        "previousActual": 500, "monthPlan": 100, "monthActual": 42}]}
        self.acc = MonthlyServiceActuals(date(2026, 10, 24))
        self.acc.observe(date(2026, 10, 1), "ЖК №10", "001441", 1)
        self.meta = {"sourceStart": "2026-10-01", "sourceEnd": "2026-10-23",
                     "source": "services.xlsx", "sheet": "Лист1", "sha256": "abc"}

    def result(self):
        return snapshot_from_services(self.config, {"meta": self.meta, "serviceActuals": self.acc.export()})

    def test_sum_quantity_not_rows_preserve_previous_and_override(self):
        self.acc.observe(date(2026, 10, 2), "ЖК №10", 1983, 2)
        self.acc.observe(date(2026, 10, 23), "ЖК №10", "001984", "3")
        original = copy.deepcopy(self.config)
        value = self.result()
        self.assertEqual(value["total"]["monthActual"], 5)
        self.assertEqual(value["total"]["yearActual"], 505)
        self.assertEqual(value["total"]["monthPlan"], 99)
        self.assertEqual(value["asOf"], "2026-10-23")
        self.assertEqual(value["actualSource"]["quantityByCode"], {"001983": 2, "001984": 3})
        self.assertEqual(self.config, original)
        self.assertTrue(value["warnings"])

    def test_bounds_inclusive_and_only_target_codes(self):
        for day, code, qty in [(date(2026, 9, 30), 1983, 100),
                              (date(2026, 10, 1), 1983, 2),
                              (date(2026, 10, 24), 1984, 3),
                              (date(2026, 10, 25), 1984, 100),
                              (date(2026, 10, 2), 1441, 100)]:
            self.acc.observe(day, "ЖК №10", code, qty)
        self.meta["sourceEnd"] = "2026-10-31"
        self.assertEqual(self.result()["total"]["monthActual"], 5)
        self.assertEqual(self.result()["asOf"], "2026-10-24")

    def test_real_zero_requires_department_presence(self):
        self.assertEqual(self.result()["total"]["monthActual"], 0)

    def test_missing_department_is_not_zero(self):
        self.acc.present.clear()
        with self.assertRaisesRegex(AnnualPlanError, "отсутствуют"):
            self.result()

    def test_unmatched_target_is_not_discarded(self):
        self.acc.observe(date(2026, 10, 2), "Неизвестное отделение", 1984, 3)
        with self.assertRaisesRegex(AnnualPlanError, "вне подразделений"):
            self.result()

    def test_non_target_unknown_department_is_ignored(self):
        self.acc.observe(date(2026, 10, 2), "Лаборатория", 9999, 1)
        self.assertEqual(self.result()["total"]["monthActual"], 0)

    def test_merger_and_name_alias(self):
        self.config["departments"][0]["name"] = "ЦЖЗ на Планетной"
        self.acc.observe(date(2026, 10, 3), "Женская консультация №5", "001983", 7)
        self.assertEqual(self.result()["total"]["monthActual"], 7)

    def test_ambiguous_plan_department_rejected(self):
        self.config["departments"].append({**self.config["departments"][0], "name": "Женская консультация №10"})
        with self.assertRaisesRegex(AnnualPlanError, "Два подразделения"):
            self.result()

    def test_wrong_month_or_year_rejected(self):
        for value in ["2026-09-22", "2025-10-22"]:
            self.config["asOf"] = value
            with self.subTest(value=value), self.assertRaisesRegex(AnnualPlanError, "не совпадает"):
                self.result()

    def test_partial_source_rejected(self):
        self.meta["sourceStart"] = "2026-10-05"
        with self.assertRaisesRegex(AnnualPlanError, "с 1-го числа"):
            self.result()

    def test_bad_quantities_rejected(self):
        for value in [-1, 1.5, "NaN", "Infinity", "", True, 1000000001]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                self.acc.observe(date(2026, 10, 2), "ЖК №10", "001983", value)

    def test_manual_snapshot_stays_unchanged(self):
        self.assertEqual(snapshot(self.config)["total"]["monthActual"], 42)
        self.assertEqual(snapshot(self.config)["asOf"], "2026-10-22")

    def test_missing_accumulator_rejected(self):
        with self.assertRaisesRegex(AnnualPlanError, "Сначала прочитайте"):
            snapshot_from_services(self.config, {"meta": self.meta})

    def test_duplicate_service_rows_are_added(self):
        for _ in range(2):
            self.acc.observe(date(2026, 10, 2), "ЖК №10", "001983", 3)
        self.assertEqual(self.result()["total"]["monthActual"], 6)

    def test_no_source_data_in_month(self):
        self.meta["sourceEnd"] = "2026-09-30"
        with self.assertRaisesRegex(AnnualPlanError, "нет данных"):
            self.result()


class CollectionIntegrationTests(unittest.TestCase):
    class Sheet:
        title = "Услуги"
        rows = [
            ["Дата", "Код услуги", "Кол-во", "Тариф", "Врач", "Подразделение"],
            [date(2026, 10, 1), 1983, 2, 10, "Тестовый врач", "ЖК №10"],
            [date(2026, 10, 23), "001984", 3, 10, "Тестовый врач", "ЖК №10"],
        ]
        def iter_values(self, min_row=1, max_row=None):
            return iter(self.rows[min_row-1:max_row])

    @contextmanager
    def source(self, path):
        yield [self.Sheet()], None

    def test_single_read_full_month_even_when_chart_starts_later(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "source.xlsx"
            path.touch()
            with patch.object(dashboard, "_open_source", side_effect=self.source) as reader:
                data = dashboard.collect(path, date(2026, 10, 20), date(2026, 10, 24), actuals_from_services=True)
            self.assertEqual(reader.call_count, 1)
            self.assertEqual(data["meta"]["totalQuantity"], 3)
            self.assertEqual(sum(r["quantity"] for r in data["serviceActuals"]["rows"]), 5)
            self.assertEqual(data["meta"]["sourceStart"], "2026-10-01")

    def test_manual_collect_has_no_automatic_payload(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "source.xlsx"
            path.touch()
            with patch.object(dashboard, "_open_source", side_effect=self.source):
                data = dashboard.collect(path, date(2026, 10, 1), date(2026, 10, 24))
            self.assertNotIn("serviceActuals", data)

    def test_invalid_mode_does_not_overwrite_html(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "report.html"
            path.write_text("keep", encoding="utf-8")
            with self.assertRaises(dashboard.DashboardError):
                dashboard.write_dashboard({}, path, overwrite=True, include_plan=False, actuals_from_services=True)
            self.assertEqual(path.read_text(encoding="utf-8"), "keep")


if __name__ == "__main__":
    unittest.main()
