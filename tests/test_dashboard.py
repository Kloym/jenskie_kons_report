import hashlib
import json
import tempfile
import unittest
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from openpyxl import Workbook
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from zhk_dashboard import collect, write_dashboard, latest_snapshot, DashboardError

HEADERS = ["Дата", "Код услуги", "Наименование услуги", "Кол-во", "Тариф", "Врач", "Подразделение", "ФИО"]
DAY = date(2026, 10, 2)

class DashboardTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
    def tearDown(self):
        self.temp.cleanup()
    def book(self, rows):
        p = self.root / "input.xlsx"
        w = Workbook(); s = w.active; s.append(HEADERS)
        for row in rows: s.append(row)
        w.save(p); w.close()
        return p
    def row(self, code=1983, quantity=2, tariff=100.25, doctor="Тестовый Врач", day=DAY):
        return [day, code, "Услуга", quantity, tariff, doctor, "ЖК №5", "СЕКРЕТНЫЙ ПАЦИЕНТ"]
    def read(self, p):
        return collect(p, date(2026,10,1), date(2026,10,23))
    def test_price_times_quantity_and_no_patient_data(self):
        p = self.book([self.row(),self.row(code="001984",quantity=3,tariff=10)])
        before=hashlib.sha256(p.read_bytes()).hexdigest()
        d=self.read(p)
        self.assertEqual(d["meta"]["totalQuantity"],5)
        self.assertEqual(d["meta"]["totalCents"],23050)
        self.assertEqual(d["services"][0]["code"],"001983")
        self.assertEqual(d["departments"],["ЦЖЗ на Планетной"])
        self.assertNotIn("СЕКРЕТНЫЙ",json.dumps(d,ensure_ascii=False))
        self.assertEqual(hashlib.sha256(p.read_bytes()).hexdigest(),before)
    def test_row_rounding_and_duplicates_retained(self):
        d=self.read(self.book([self.row(quantity=3,tariff=.005),self.row(quantity=3,tariff=.005)]))
        self.assertEqual(d["meta"]["totalQuantity"],6)
        self.assertEqual(d["meta"]["totalCents"],4)
        self.assertEqual(d["records"][0][6],2)
    def test_missing_price_uses_unique_unit_price(self):
        d=self.read(self.book([self.row(quantity=2,tariff=100),self.row(quantity=3,tariff=None)]))
        self.assertEqual(d["meta"]["totalCents"],50000)
        self.assertEqual(d["meta"]["stats"]["restored_tariff_rows"],1)
    def test_zero_price_not_replaced(self):
        d=self.read(self.book([self.row(tariff=100),self.row(quantity=3,tariff=0)]))
        self.assertEqual(d["meta"]["totalCents"],20000)
        self.assertEqual(d["meta"]["stats"]["zero_tariff_rows"],1)
    def test_conflicting_price_fails(self):
        p=self.book([self.row(tariff=10),self.row(tariff=20),self.row(tariff=None)])
        with self.assertRaises(DashboardError):self.read(p)
    def test_invalid_and_missing_quantity_fails(self):
        for q in (None,-1,1.5,"bad"):
            with self.subTest(q=q),self.assertRaises(DashboardError):
                self.read(self.book([self.row(quantity=q)]))
    def test_missing_doctor_is_visible_not_dropped(self):
        d=self.read(self.book([self.row(doctor=None)]))
        self.assertEqual(d["meta"]["totalQuantity"],2)
        self.assertEqual(d["doctors"],["Врач не указан"])
    def test_date_coverage_and_inclusive_end(self):
        d=self.read(self.book([self.row(day=date(2026,9,30)),self.row(day=date(2026,10,22))]))
        self.assertEqual(d["meta"]["totalQuantity"],2)
        self.assertEqual(d["meta"]["dataEnd"],"2026-10-22")
        self.assertNotIn("2026-10-23",d["coverage"])
        self.assertEqual(d["meta"]["stats"]["outside_period"],1)
    def test_html_escape_and_no_unintended_template_replacement(self):
        d=self.read(self.book([self.row(doctor="</script><script>alert(1)</script>{{SCRIPT}}")]))
        output=write_dashboard(d,self.root/"report.html")
        text=output.read_text()
        self.assertIn("\\u003c/script\\u003e",text)
        self.assertNotIn("</script><script>alert(1)",text)
        self.assertIn("{{SCRIPT}}",text)
        self.assertIn("Anime.js - UMD minified bundle",text)
        self.assertNotIn("{{ANIME}}",text)
        with self.assertRaises(DashboardError):write_dashboard(d,output)
    def test_multiple_sheets_require_explicit_selection(self):
        p=self.book([self.row()])
        from openpyxl import load_workbook
        w=load_workbook(p);s=w.create_sheet("second");s.append(HEADERS);s.append(self.row());w.save(p);w.close()
        with self.assertRaises(DashboardError):self.read(p)
        d=collect(p,date(2026,10,1),date(2026,10,23),sheet_name="second")
        self.assertEqual(d["meta"]["totalQuantity"],2)
    def test_latest_snapshot_not_combined(self):
        p=self.book([self.row()]);q=self.root/"other.xlsx";q.write_bytes(p.read_bytes())
        import os
        os.utime(p,(1,1));os.utime(q,(2,2))
        lock=self.root/"~$lock.xlsx";lock.write_bytes(b"")
        self.assertEqual(latest_snapshot(self.root),q)

if __name__=="__main__":unittest.main()
