"""Автономный HTML-отчёт по услугам ЖК. Python 3.10+.
Без параметров открывается окно выбора Excel и сохранения HTML.
Правило выгрузки подтверждено пользователем: Тариф = цена за одну услугу.
Входные файлы не меняются; в HTML попадают только агрегаты без пациентов.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
from collections import Counter, defaultdict
from datetime import date, datetime, timezone, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from zhk_monthly_report import (
    _open_source, clean_text, decimal_number, identify_department,
    normalize_key, parse_date, service_code, ensure_dependencies,
)

from zhk_annual_plan import load_config, snapshot, edit_plan

BASE = Path(__file__).resolve().parent
TARGET_CODES = ("001983", "001984")
CENT = Decimal("0.01")
ALIASES = {
    "date": {"дата", "дата услуги", "дата оказания услуги", "дата выполнения"},
    "code": {"код услуги", "код мед услуги", "код медуслуги"},
    "name": {"наименование услуги", "название услуги", "услуга"},
    "quantity": {"кол во", "количество", "количество услуг"},
    "tariff": {"тариф", "тариф руб", "цена", "цена руб", "цена тариф руб"},
    "doctor": {"врач", "фио врача", "исполнитель", "врач исполнитель"},
    "department": {"подразделение", "наименование подразделения", "отделение", "жк"},
}
REQUIRED = ("date", "code", "quantity", "tariff", "doctor", "department")


class DashboardError(ValueError):
    pass


def local_today():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Europe/Moscow")).date()
    except (ImportError, KeyError):
        return date.today()


def user_date(value):
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    raise DashboardError("Дата должна быть в формате ДД.ММ.ГГГГ или ГГГГ-ММ-ДД.")


def header_of(sheet):
    for number, row in enumerate(list(sheet.iter_values(max_row=30)), 1):
        keys = [normalize_key(v) for v in row]
        cols = {}
        for field, names in ALIASES.items():
            hits = [i for i, key in enumerate(keys) if key in names]
            if len(hits) == 1:
                cols[field] = hits[0]
        if all(field in cols for field in REQUIRED):
            return number, cols
    return None


def collect(path, start, end, *, sheet_name=None, all_departments=False,
            restore_missing=True, progress=None):
    """One full snapshot, Decimal row totals, compact daily aggregates."""
    path = Path(path).resolve()
    if start > end:
        raise DashboardError("Начало периода позже окончания.")
    if (end - start).days > 370:
        raise DashboardError("Выберите период не длиннее 371 дня.")
    if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
        raise DashboardError("Нужен существующий файл .xlsx, .xls или .xlsm.")
    ensure_dependencies(path.suffix.lower() == ".xls")
    stats = Counter()
    groups = defaultdict(lambda: [Decimal(0), Decimal(0), 0, 0])
    doctors, names = {}, {}
    donors, pending = defaultdict(set), []
    coverage = set()
    observed_start = observed_end = None

    def add(key, quantity, price, restored=False):
        amount = (price * quantity).quantize(CENT, rounding=ROUND_HALF_UP)
        group = groups[key]
        group[0] += quantity
        group[1] += amount
        group[2] += 1
        group[3] += int(restored)
        stats["included_rows"] += 1

    with _open_source(path) as (sheets, epoch):
        candidates = []
        for sheet in sheets:
            if sheet_name and sheet.title != sheet_name:
                continue
            found = header_of(sheet)
            if found:
                candidates.append((sheet, found))
        if not candidates:
            raise DashboardError("Не найдены однозначные колонки: дата, код услуги, "
                                 "количество, тариф, врач и подразделение.")
        if len(candidates) != 1:
            raise DashboardError("Несколько подходящих листов. Укажите --sheet с именем листа.")
        sheet, (header_row, cols) = candidates[0]
        for row_no, row in enumerate(sheet.iter_values(min_row=header_row + 1), header_row + 1):
            if not any(clean_text(v) for v in row):
                stats["blank_rows"] += 1
                continue
            stats["source_rows"] += 1
            where = f"Лист «{sheet.title}», строка {row_no}"
            def cell(field):
                index = cols.get(field)
                return row[index] if index is not None and index < len(row) else None
            try:
                day = parse_date(cell("date"), epoch)
            except (ValueError, TypeError) as exc:
                raise DashboardError(f"{where}: неверная дата. Отчёт не сформирован.") from exc
            observed_start = min(observed_start or day, day)
            observed_end = max(observed_end or day, day)
            if not start <= day <= end:
                stats["outside_period"] += 1
                continue
            raw_department = clean_text(cell("department"))
            department = identify_department(raw_department, day)
            if not department and not all_departments:
                stats["excluded_departments"] += 1
                continue
            department = department or raw_department or "Подразделение не указано"
            try:
                code = service_code(cell("code"))
                if not code:
                    raise ValueError("Не заполнен код услуги")
                quantity = decimal_number(cell("quantity"), "Количество")
                if quantity < 0 or quantity != quantity.to_integral_value():
                    raise ValueError("Количество должно быть целым и неотрицательным")
                if quantity > 10**9:
                    raise ValueError("Количество превышает допустимый предел")
                raw_price = cell("tariff")
                price = None if not clean_text(raw_price) else decimal_number(raw_price, "Тариф")
                if price is not None and price < 0:
                    raise ValueError("Тариф должен быть неотрицательным")
            except (ValueError, TypeError) as exc:
                raise DashboardError(f"{where}: {exc}. Отчёт не сформирован.") from exc
            doctor = clean_text(cell("doctor")) or "Врач не указан"
            if doctor == "Врач не указан":
                stats["missing_doctor_rows"] += 1
            doctor_key = normalize_key(doctor)
            doctors.setdefault(doctor_key, doctor)
            name = clean_text(cell("name")) or code
            names.setdefault(code, name)
            if not clean_text(cell("name")):
                stats["missing_name_rows"] += 1
            key = (day.isoformat(), department, doctor_key, code)
            donor_key = (day.year, day.month, code)
            coverage.add(day.isoformat())
            if price is None:
                pending.append((key, quantity, donor_key, row_no))
            else:
                if price > 0:
                    donors[donor_key].add(price)
                elif quantity > 0:
                    stats["zero_tariff_rows"] += 1
                add(key, quantity, price)
            if progress and stats["source_rows"] % 10000 == 0:
                progress(f"Прочитано {stats['source_rows']:,} строк…".replace(",", " "))
        selected_sheet = sheet.title

    for key, quantity, donor_key, row_no in pending:
        choices = donors[donor_key]
        if quantity == 0:
            price = Decimal(0)
        elif restore_missing and len(choices) == 1:
            price = next(iter(choices))
        else:
            raise DashboardError(f"Строка {row_no}: пустой тариф; нет единственной цены "
                                 "этого кода в том же месяце. Заполните тариф в выгрузке.")
        add(key, quantity, price, True)
        stats["restored_tariff_rows"] += 1
    if not groups:
        raise DashboardError("За выбранный период нет услуг выбранных подразделений.")
    ds = sorted({k[1] for k in groups})
    docs = sorted(doctors, key=lambda k: doctors[k])
    codes = sorted(names)
    di, dr, ci = ({v: i for i, v in enumerate(values)} for values in (ds, docs, codes))
    records = []
    for (day, department, doctor, code), (quantity, amount, rows, restored) in sorted(groups.items()):
        cents = int(amount * 100)
        if abs(cents) > 9_007_199_254_740_991:
            raise DashboardError("Сумма превышает точный диапазон HTML-отчёта.")
        records.append([day, di[department], dr[doctor], ci[code], int(quantity), cents, rows, restored])
    total_q = sum(r[4] for r in records)
    total_cents = sum(r[5] for r in records)
    if total_q > 9_007_199_254_740_991 or total_cents > 9_007_199_254_740_991:
        raise DashboardError("Итог превышает точный диапазон HTML-отчёта.")
    with path.open("rb") as handle:
        digest = hashlib.sha256()
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return {
        "version": 1,
        "meta": {
            "source": path.name, "sheet": selected_sheet, "sha256": digest.hexdigest(),
            "start": start.isoformat(), "end": end.isoformat(),
            "dataStart": min(coverage), "dataEnd": max(coverage),
            "sourceStart": observed_start.isoformat(), "sourceEnd": observed_end.isoformat(),
            "created": datetime.now(timezone(timedelta(hours=3))).isoformat(timespec="seconds"),
            "today": local_today().isoformat(), "amountRule": "unit_price",
            "stats": dict(stats), "totalQuantity": total_q, "totalCents": total_cents,
            "scope": "Все подразделения" if all_departments else "Женские консультации",
        },
        "departments": ds, "doctors": [doctors[k] for k in docs],
        "services": [{"code": code, "name": names[code]} for code in codes],
        "coverage": sorted(coverage), "records": records, "targets": list(TARGET_CODES),
    }


def write_dashboard(data, output, *, overwrite=False, source=None, plan_path=None, include_plan=True):
    output = Path(output).resolve()
    if output.suffix.lower() != ".html":
        raise DashboardError("Выходной файл должен иметь расширение .html.")
    if source and output == Path(source).resolve():
        raise DashboardError("Нельзя перезаписывать исходный файл.")
    if output.exists() and not overwrite:
        raise DashboardError(f"Файл уже существует: {output}. Укажите --overwrite.")
    data = dict(data)
    data.pop("annualPlan", None)
    if include_plan:
        plan_source = Path(plan_path) if plan_path is not None else BASE / "annual_plan.json"
        if plan_source.exists() or plan_path is not None:
            data["annualPlan"] = snapshot(load_config(plan_source))
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    # Neutralize script terminators and HTML parser delimiters in source labels.
    payload = payload.replace("&", "\\u0026").replace("<", "\\u003c").replace(">", "\\u003e")
    payload = payload.replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    template = (BASE / "dashboard_assets" / "dashboard.html").read_text(encoding="utf-8")
    css = (BASE / "dashboard_assets" / "dashboard.css").read_text(encoding="utf-8")
    css += "\n" + (BASE / "dashboard_assets" / "annual_plan.css").read_text(encoding="utf-8")
    css += "\n" + (BASE / "dashboard_assets" / "clinic_pace.css").read_text(encoding="utf-8")
    js = (BASE / "dashboard_assets" / "annual_plan.js").read_text(encoding="utf-8")
    js += "\n" + (BASE / "dashboard_assets" / "daily_plan_core.js").read_text(encoding="utf-8")
    js += "\n" + (BASE / "dashboard_assets" / "clinic_pace.js").read_text(encoding="utf-8")
    js += "\n" + (BASE / "dashboard_assets" / "dashboard.js").read_text(encoding="utf-8")
    anime = (BASE / "dashboard_assets" / "vendor" / "anime-4.5.0.umd.min.js").read_text(encoding="utf-8")
    anime_license = (BASE / "dashboard_assets" / "vendor" / "LICENSE-anime.md").read_text(encoding="utf-8")
    anime = "/*\n" + anime_license + "\n*/\n" + anime
    # Single pass: source values containing template tokens remain literal.
    import re
    values = {"/* {{STYLE}} */": css, "{{STYLE}}": css, "{{SCRIPT}}": js,
              "{{ANIME}}": anime, "{{DATA}}": payload}
    pattern = "|".join(re.escape(token) for token in values)
    html = re.sub(pattern, lambda match: values[match[0]], template)
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".dashboard-", suffix=".tmp", dir=output.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(html)
        os.replace(temporary, output)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return output


def latest_snapshot(folder):
    files = [p for p in Path(folder).glob("*") if p.is_file()
             and p.suffix.lower() in {".xls", ".xlsx", ".xlsm"} and not p.name.startswith("~$")]
    if not files:
        raise DashboardError(f"Положите выгрузку в {folder} или укажите --input.")
    return max(files, key=lambda p: (p.stat().st_mtime_ns, p.name))


def launch_dashboard():
    """Small desktop launcher; XLSX work stays outside the Tk event loop."""
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    import webbrowser

    root = tk.Tk()
    root.title("Ритм консультаций • HTML-дашборд")
    root.geometry("700x430")
    root.configure(padx=24, pady=22)
    path_var, status = tk.StringVar(), tk.StringVar(value="Выберите полную выгрузку услуг за месяц.")
    today = local_today()
    from_var, to_var = tk.StringVar(value=today.replace(day=1).strftime("%d.%m.%Y")), tk.StringVar(value=today.strftime("%d.%m.%Y"))
    ttk.Label(root, text="Ритм консультаций", font=("", 20)).pack(anchor="w")
    ttk.Label(root, text="Количество, суммы и ежедневная динамика врачей").pack(anchor="w", pady=(4, 22))
    row = ttk.Frame(root); row.pack(fill="x")
    ttk.Entry(row, textvariable=path_var).pack(side="left", fill="x", expand=True)
    def choose():
        selected = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if selected: path_var.set(selected)
    ttk.Button(row, text="Выбрать Excel", command=choose).pack(side="right", padx=(8, 0))
    dates = ttk.Frame(root); dates.pack(fill="x", pady=18)
    for label, var in (("С", from_var), ("По", to_var)):
        ttk.Label(dates, text=label).pack(side="left", padx=(0, 8))
        ttk.Entry(dates, textvariable=var, width=16).pack(side="left", padx=(0, 20))
    plan_row = ttk.Frame(root); plan_row.pack(fill="x", pady=(0, 14))
    include_plan = tk.BooleanVar(value=True)
    ttk.Checkbutton(plan_row, text="Добавить годовой план", variable=include_plan).pack(side="left")
    plan_status = tk.StringVar()
    def refresh_plan(saved=None):
        try:
            current = saved or (snapshot(load_config(BASE / "annual_plan.json")) if (BASE / "annual_plan.json").exists() else None)
            plan_status.set("На " + date.fromisoformat(current["asOf"]).strftime("%d.%m.%Y") if current else "Ещё не заполнен")
        except ValueError:
            plan_status.set("Проверьте сохранённый план")
    ttk.Label(plan_row, textvariable=plan_status).pack(side="left", padx=10)
    ttk.Button(plan_row, text="Годовой план…", command=lambda: edit_plan(root, BASE / "annual_plan.json", refresh_plan)).pack(side="right")
    refresh_plan()
    ttk.Label(root, textvariable=status, wraplength=600).pack(anchor="w", pady=(0, 16))
    messages = queue.Queue()
    def build():
        try:
            source = Path(path_var.get())
            start, end = user_date(from_var.get()), user_date(to_var.get())
            if not source.is_file(): raise DashboardError("Сначала выберите файл Excel.")
        except ValueError as exc:
            messagebox.showerror("Проверьте данные", str(exc)); return
        output = filedialog.asksaveasfilename(defaultextension=".html", initialfile=f"dashboard_{end:%Y-%m-%d}.html",
                                            filetypes=[("HTML", "*.html")])
        if not output: return
        use_plan = include_plan.get()
        button.configure(state="disabled")
        status.set("Чтение Excel…")
        def worker():
            try:
                data = collect(source, start, end, progress=lambda text: messages.put(("progress", text)))
                result = write_dashboard(data, output, overwrite=True, source=source, include_plan=use_plan)
                messages.put(("done", result))
            except Exception as exc:
                messages.put(("error", str(exc)))
        threading.Thread(target=worker, daemon=True).start()
    def poll():
        try:
            while True:
                kind, value = messages.get_nowait()
                if kind == "progress": status.set(value)
                else:
                    button.configure(state="normal")
                    if kind == "done":
                        status.set(f"Готово: {value.name}"); webbrowser.open(value.as_uri())
                    else:
                        status.set("Отчёт не сформирован."); messagebox.showerror("Ошибка", value)
        except queue.Empty: pass
        root.after(120, poll)
    button = ttk.Button(root, text="Создать HTML-дашборд", command=build); button.pack(anchor="w")
    poll(); root.mainloop()


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, help="Полная выгрузка месяца, один файл")
    parser.add_argument("--input-dir", type=Path, default=BASE / "incoming", help="Берётся только самый новый файл")
    parser.add_argument("--date-from", type=user_date)
    parser.add_argument("--date-to", type=user_date, default=local_today())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--sheet")
    parser.add_argument("--all-departments", action="store_true")
    parser.add_argument("--no-restore-missing", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--gui", action="store_true")
    parser.add_argument("--annual-plan", type=Path, help="Сохранённый годовой план JSON; по умолчанию рядом с программой")
    parser.add_argument("--without-annual-plan", action="store_true", help="Собрать отчёт без блока годового плана")
    parser.add_argument("--edit-plan", action="store_true", help="Открыть окно ручного ввода годового плана")
    args = parser.parse_args(argv)
    if args.edit_plan:
        import tkinter as tk
        root = tk.Tk(); root.withdraw()
        editor = edit_plan(root, args.annual_plan or BASE / "annual_plan.json")
        if editor is not None:
            root.wait_window(editor)
        root.destroy(); return 0
    if args.gui or not argv:
        launch_dashboard(); return 0
    try:
        source = args.input or latest_snapshot(args.input_dir)
        start = args.date_from or args.date_to.replace(day=1)
        output = args.output or BASE / "dashboards" / f"dashboard_{start}_{args.date_to}.html"
        data = collect(source, start, args.date_to, sheet_name=args.sheet,
                       all_departments=args.all_departments, restore_missing=not args.no_restore_missing)
        result = write_dashboard(data, output, overwrite=args.overwrite, source=source,
                                 plan_path=args.annual_plan, include_plan=not args.without_annual_plan)
        print(f"Готово: {result}\nУслуг: {data['meta']['totalQuantity']:,}; "
              f"сумма: {data['meta']['totalCents'] / 100:,.2f} руб.\n"
              f"Последний день с данными: {data['meta']['dataEnd']}")
        return 0
    except (DashboardError, OSError, ValueError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr); return 2


if __name__ == "__main__":
    raise SystemExit(main())
