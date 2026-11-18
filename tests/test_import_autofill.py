"""Service lookup and real Tk editor behavior (synthetic data only)."""
import hashlib
import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import zhk_dashboard as dashboard
from zhk_import_review import ImportReview, ImportReviewDialog, line_amount

FIRST, LAST = date(2026, 11, 1), date(2026, 11, 17)
NAME = "Диспансеризация женщин 18–29 лет"
OTHER_NAME = "Диспансеризация женщин 30–49 лет"


def record(**changes):
    values = dict(date=FIRST, code="001983", name=NAME, quantity=2,
                  tariff="3446,40", doctor="Тестовый врач", department="ЖК №10")
    values.update(changes)
    return values


def review_of(*donors, restore_missing=True):
    rows = list(enumerate(donors or [record()], 2))
    rows.append((65, record(code=None, name="", quantity="", tariff="")))
    review = ImportReview("Услуги", CALENDAR_WINDOWS_1900, rows, FIRST, LAST,
                          restore_missing=restore_missing)
    review.evaluate()
    return review


class CatalogTests(unittest.TestCase):
    def test_normalized_code_and_unit_price_not_row_total(self):
        r = review_of(record(code=1983), record(code="'001983"), record(code=1983.0))
        choices = r.suggestions(65, record(code="1983"))
        self.assertEqual(choices.code, "001983")
        self.assertEqual(choices.names, (NAME,))
        self.assertEqual(choices.tariffs, (Decimal("3446.40"),))
        self.assertEqual(sorted(r.catalog.codes), ["001983"])

    def test_tariff_stays_in_same_year_and_month(self):
        r = review_of(record(), record(date=date(2026, 10, 1), tariff=100, name="Старое название"),
                      record(date=date(2025, 11, 1), tariff=200))
        choices = r.suggestions(65, record())
        self.assertEqual(choices.tariffs, (Decimal("3446.40"),))
        self.assertEqual(choices.names, (NAME,))
        self.assertEqual(r.suggestions(65, record(date=date(2026, 10, 2))).tariffs, (Decimal(100),))
        self.assertEqual(r.suggestions(65, record(date=date(2026, 12, 2))).tariffs, ())
        self.assertEqual(r.suggestions(65, record(date="bad")).tariffs, ())

    def test_name_can_come_from_another_month_but_price_cannot(self):
        r = review_of(record(date=date(2026, 10, 1)))
        choices = r.suggestions(65, record())
        self.assertEqual(choices.names, (NAME,))
        self.assertEqual(choices.tariffs, ())

    def test_edit_exclude_and_restore_refresh_donors(self):
        r = review_of(record(tariff=100), record(tariff=100))
        r.edit(2, {**r.values(2), "tariff": "200"})
        self.assertEqual(r.suggestions(65, record()).tariffs, (Decimal(100), Decimal(200)))
        r.exclude([3])
        self.assertEqual(r.suggestions(65, record()).tariffs, (Decimal(200),))
        r.restore([2, 3])
        self.assertEqual(r.suggestions(65, record()).tariffs, (Decimal(100),))
        r.exclude([2, 3])
        self.assertEqual(r.suggestions(65, record()).tariffs, ())

    def test_row_never_donates_to_itself(self):
        r = review_of(record())
        self.assertEqual(r.suggestions(2, r.values(2)).tariffs, ())
        self.assertEqual(r.suggestions(2, r.values(2)).names, ())
        r.edit(65, record(tariff=50))
        r.exclude([2])
        self.assertEqual(r.suggestions(65, r.values(65)).tariffs, ())

    def test_invalid_and_zero_prices_are_not_automatic_donors(self):
        r = review_of(record(tariff=0), record(tariff=-1), record(tariff="NaN"),
                      record(quantity=-1), record(quantity=""), record(quantity="1.5"),
                      record(department=""), record(date="bad"))
        self.assertEqual(r.suggestions(65, record()).tariffs, ())

    def test_amount_uses_quantity_and_rounds_once_to_kopecks(self):
        self.assertEqual(line_amount("2", "3 446,40"), Decimal("6892.80"))
        self.assertEqual(line_amount("3", "0,005"), Decimal("0.02"))
        self.assertEqual(line_amount("0", "100"), Decimal("0.00"))
        self.assertEqual(line_amount("2", "0"), Decimal("0.00"))
        for quantity, price in [("", "100"), ("2.5", "100"), ("-1", "100"),
                                ("2", ""), ("2", "-1"), ("2", "Infinity"),
                                ("1000000001", "1")]:
            with self.subTest(quantity=quantity, price=price), self.assertRaises(ValueError):
                line_amount(quantity, price)


class DialogTests(unittest.TestCase):
    def setUp(self):
        import tkinter as tk
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest("Tk display unavailable: " + str(exc))
        self.root.withdraw()
        self.callback_errors = []
        self.root.report_callback_exception = lambda *args: self.callback_errors.append(args)
        self.dialog = None

    def tearDown(self):
        if hasattr(self, "root"):
            self.root.destroy()
            self.assertEqual(self.callback_errors, [], "Unexpected Tk callback failure")

    def open_review(self, review):
        self.dialog = ImportReviewDialog(self.root, review, lambda accepted: None)
        self.dialog.window.withdraw()
        self.root.update()
        self.assertEqual(self.dialog.current, 65)
        return self.dialog

    def choose(self, code):
        self.dialog.fields["code"].set(code)
        self.dialog.entries["code"].event_generate("<<ComboboxSelected>>")
        self.root.update()

    def test_selection_fills_fields_and_quantity_updates_amount(self):
        r = review_of()
        d = self.open_review(r)
        self.choose("001983")
        self.assertEqual(d.fields["name"].get(), NAME)
        self.assertEqual(Decimal(d.fields["tariff"].get().replace(",", ".")), Decimal("3446.40"))
        self.assertEqual(d.fields["quantity"].get(), "")
        self.assertEqual(d.amount_text.get(), "Сумма строки: —")
        d.fields["quantity"].set("2")
        self.assertEqual(d.amount_text.get(), "Сумма строки: 6 892,80 ₽")
        d.fields["quantity"].set("3")
        self.assertEqual(d.amount_text.get(), "Сумма строки: 10 339,20 ₽")
        self.assertNotIn("Не заполнен код услуги", d.explanation.get())
        d.save()
        self.root.update()
        self.assertEqual(r.values(65)["quantity"], "3")
        self.assertEqual(r.values(65)["name"], NAME)
        self.assertTrue(r.evaluate())
        self.assertEqual(r.prepared[-1].price, Decimal("3446.40"))

    def test_changing_service_replaces_previous_name_and_price(self):
        d = self.open_review(review_of(record(), record(code="001984", name=OTHER_NAME, tariff="2109,82")))
        self.choose("001983")
        d.fields["quantity"].set("2")
        self.choose("001984")
        self.assertEqual(d.fields["name"].get(), OTHER_NAME)
        self.assertEqual(d.fields["tariff"].get(), "2109,82")
        self.assertEqual(d.amount_text.get(), "Сумма строки: 4 219,64 ₽")
        self.choose("999999")
        self.assertEqual(d.fields["name"].get(), "")
        self.assertEqual(d.fields["tariff"].get(), "")
        self.assertEqual(d.amount_text.get(), "Сумма строки: —")

    def test_manual_price_and_zero_survive_quantity_changes(self):
        d = self.open_review(review_of())
        self.choose("001983")
        d.fields["tariff"].set("100,25")
        d.fields["quantity"].set("2")
        self.choose("001983")
        self.assertEqual(d.fields["tariff"].get(), "100,25")
        self.assertEqual(d.amount_text.get(), "Сумма строки: 200,50 ₽")
        d.fields["tariff"].set("0")
        d.fields["quantity"].set("3")
        self.choose("001983")
        self.assertEqual(d.fields["tariff"].get(), "0")
        self.assertEqual(d.amount_text.get(), "Сумма строки: 0,00 ₽")

    def test_multiple_prices_and_names_require_choice(self):
        d = self.open_review(review_of(record(tariff=100), record(tariff=200, name="Второе название")))
        self.choose("001983")
        self.assertEqual(d.fields["tariff"].get(), "")
        self.assertEqual(d.fields["name"].get(), "")
        self.assertEqual(tuple(d.entries["tariff"]["values"]), ("100", "200"))
        self.assertEqual(set(d.entries["name"]["values"]), {NAME, "Второе название"})
        d.fields["name"].set(NAME)
        d.fields["tariff"].set("200")
        d.fields["quantity"].set("2")
        self.assertEqual(d.amount_text.get(), "Сумма строки: 400,00 ₽")

    def test_month_change_reselects_tariff_and_missing_month_clears_it(self):
        d = self.open_review(review_of(record(), record(date=date(2026, 10, 1), tariff=500)))
        self.choose("001983")
        d.fields["quantity"].set("2")
        d.fields["date"].set("02.10.2026")
        d.entries["date"].event_generate("<FocusOut>")
        self.assertEqual(d.fields["tariff"].get(), "500")
        self.assertEqual(d.amount_text.get(), "Сумма строки: 1 000,00 ₽")
        d.fields["date"].set("02.10.2025")
        d.entries["date"].event_generate("<FocusOut>")
        self.assertEqual(d.fields["tariff"].get(), "")
        d.fields["date"].set("bad")
        d.entries["date"].event_generate("<FocusOut>")
        self.assertEqual(d.amount_text.get(), "Сумма строки: —")

    def test_disabled_restoration_preserves_manual_choice(self):
        d = self.open_review(review_of(restore_missing=False))
        self.choose("001983")
        self.assertEqual(d.fields["name"].get(), NAME)
        self.assertEqual(d.fields["tariff"].get(), "")
        self.assertEqual(tuple(d.entries["tariff"]["values"]), ("3446,40",))

    def test_restore_returns_original_fields_without_refilling(self):
        r = review_of()
        d = self.open_review(r)
        self.choose("001983")
        d.fields["quantity"].set("2")
        d.save()
        self.root.update()
        d.restore()
        self.root.update()
        self.assertEqual(d.fields["code"].get(), "")
        self.assertEqual(d.fields["name"].get(), "")
        self.assertEqual(d.fields["tariff"].get(), "")
        self.assertEqual(r.overrides, {})

    def test_typed_code_commits_before_apply_and_row_switch(self):
        r = review_of()
        r.review_rows.add(2)
        d = ImportReviewDialog(self.root, r, lambda accepted: None)
        self.dialog = d
        d.window.withdraw()
        d.tree.selection_set("65")
        self.root.update()
        d.fields["code"].set("1983")
        d.fields["quantity"].set("2")
        d.tree.selection_set("2")
        self.root.update()
        self.assertEqual(r.values(65)["name"], NAME)
        self.assertEqual(r.values(65)["tariff"], "3446,40")
        d.tree.selection_set("65")
        self.root.update()
        self.assertEqual(d.amount_text.get(), "Сумма строки: 6 892,80 ₽")

    def test_controls_fit_compact_window(self):
        d = self.dialog = ImportReviewDialog(self.root, review_of(), lambda accepted: None)
        d.window.transient("")
        d.window.attributes("-alpha", 0.0)
        d.window.deiconify()
        self.root.update()
        d.window.geometry("920x650")
        self.root.update()
        bottom = d.window.winfo_rooty() + d.window.winfo_height()
        for widget in (d.entries["department"], d.save_button, d.continue_button):
            self.assertTrue(widget.winfo_ismapped())
            self.assertLessEqual(widget.winfo_rooty() + widget.winfo_height(), bottom)

    def test_full_report_uses_autofill_without_rewriting_excel(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Дата", "Код услуги", "Наименование услуги", "Кол-во", "Тариф", "Врач", "Подразделение"])
            for values in [record(), record(code=None, name="", quantity="", tariff="")]:
                sheet.append([values[key] for key in ("date", "code", "name", "quantity", "tariff", "doctor", "department")])
            workbook.save(path)
            workbook.close()
            original = hashlib.sha256(path.read_bytes()).hexdigest()
            def resolve(review):
                d = self.dialog = ImportReviewDialog(self.root, review, lambda accepted: None)
                d.window.withdraw()
                self.root.update()
                self.choose("001983")
                d.fields["quantity"].set("3")
                d.save()
                return True
            data = dashboard.collect(path, FIRST, LAST, resolve_issues=resolve, actuals_from_services=True)
            self.assertEqual(data["meta"]["totalQuantity"], 5)
            self.assertEqual(data["meta"]["totalCents"], 1723200)
            self.assertEqual(sum(row["quantity"] for row in data["serviceActuals"]["rows"]), 5)
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), original)


if __name__ == "__main__":
    unittest.main()
