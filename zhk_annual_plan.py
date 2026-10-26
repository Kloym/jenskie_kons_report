"""Manual annual-plan snapshots, independent of the monthly service export."""
import calendar
import json
import os
import tempfile
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP, ROUND_CEILING
from pathlib import Path

FIELDS = ("yearPlan", "previousActual", "monthPlan", "monthActual")
LABELS = ("План года", "Факт до месяца", "План месяца", "Факт месяца на дату")
DEFAULT_NAMES = ("ЦЖЗ на Лобненской", "ЦЖЗ на Петрозаводской", "ЖК №10", "ЦЖЗ на Планетной")


class AnnualPlanError(ValueError):
    pass


def count(value, label):
    if isinstance(value, bool):
        raise AnnualPlanError(f"{label}: введите целое неотрицательное число.")
    text = "".join(str(value).split())
    if not text.isascii() or not text.isdigit():
        raise AnnualPlanError(f"{label}: введите целое неотрицательное число.")
    number = int(text)
    if number > 1_000_000_000:
        raise AnnualPlanError(f"{label}: число слишком велико, проверьте ввод.")
    return number


def validate_config(config):
    if not isinstance(config, dict) or type(config.get("version")) is not int or config.get("version") != 1:
        raise AnnualPlanError("Неизвестный формат годового плана.")
    try:
        as_of = date.fromisoformat(config["asOf"])
    except (KeyError, TypeError, ValueError):
        raise AnnualPlanError("Укажите корректную дату годового плана.") from None
    source_rows = config.get("departments")
    if not isinstance(source_rows, list) or not 1 <= len(source_rows) <= 100:
        raise AnnualPlanError("В годовом плане должно быть от 1 до 100 подразделений.")
    rows, names = [], set()
    for source in source_rows:
        if not isinstance(source, dict):
            raise AnnualPlanError("Некорректная строка подразделения.")
        name = " ".join(str(source.get("name", "")).split())
        if not name or len(name) > 160 or name.casefold() in names:
            raise AnnualPlanError("Названия подразделений должны быть заполнены и не повторяться.")
        names.add(name.casefold())
        row = {"name": name}
        for key, label in zip(FIELDS, LABELS):
            row[key] = count(source.get(key, ""), f"{name}, {label.lower()}")
        if as_of.month == 1 and row["previousActual"]:
            raise AnnualPlanError(f"{name}: в январе факт до начала месяца должен быть равен нулю.")
        rows.append(row)
    override = config.get("monthPlanTotal")
    if override not in (None, ""):
        override = count(override, "Общий план месяца из отчёта")
    else:
        override = None
    return {"version": 1, "asOf": as_of.isoformat(), "monthPlanTotal": override, "departments": rows}


def load_config(path):
    try:
        return validate_config(json.loads(Path(path).read_text(encoding="utf-8")))
    except (OSError, json.JSONDecodeError) as exc:
        raise AnnualPlanError(f"Не удалось прочитать годовой план: {exc}") from exc


def save_config(config, path):
    config = validate_config(config)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".annual-plan-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(config, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        if path.exists():
            # One recoverable previous revision; a failed validation never replaces it.
            path.with_suffix(".backup.json").write_bytes(path.read_bytes())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return config


def snapshot(config):
    config = validate_config(config)
    as_of = date.fromisoformat(config["asOf"])
    elapsed = Decimal(as_of.day) / calendar.monthrange(as_of.year, as_of.month)[1]
    year_fraction = (as_of.month - 1 + elapsed) / 12
    prior_fraction = Decimal(as_of.month - 1) / 12

    def rounded(value, places=2):
        return float(value.quantize(Decimal(10) ** -places, rounding=ROUND_HALF_UP))

    def ratio(actual, target):
        return Decimal(actual) / target * 100 if target else None

    def finish(row):
        row = dict(row)
        row["yearActual"] = row["previousActual"] + row["monthActual"]
        year_target = Decimal(row["yearPlan"]) * year_fraction
        month_target = Decimal(row["monthPlan"]) * elapsed
        prior_rate = ratio(row["previousActual"], Decimal(row["yearPlan"]) * prior_fraction)
        year_rate = ratio(row["yearActual"], year_target)
        for key, value in {
            "yearShare": ratio(row["yearActual"], Decimal(row["yearPlan"])),
            "priorRate": prior_rate,
            "yearRate": year_rate,
            "monthRate": ratio(row["monthActual"], month_target),
        }.items():
            row[key] = rounded(value, 1) if value is not None else None
        row["trend"] = rounded(year_rate - prior_rate) if year_rate is not None and prior_rate is not None else None
        row["gapToDate"] = int(year_target.to_integral_value(rounding=ROUND_CEILING)) - row["yearActual"]
        row["yearTargetToDate"] = rounded(year_target)
        row["monthTargetToDate"] = rounded(month_target)
        row["remainingYear"] = max(0, row["yearPlan"] - row["yearActual"])
        return row

    totals = {key: sum(row[key] for row in config["departments"]) for key in FIELDS}
    warnings = []
    override = config["monthPlanTotal"]
    if override is not None:
        totals["monthPlan"] = override
    return {
        "asOf": config["asOf"], "year": as_of.year,
        "yearFraction": float(year_fraction), "monthFraction": float(elapsed),
        "total": finish({"name": "Всё учреждение", **totals}),
        "departments": [finish(row) for row in config["departments"]],
        "warnings": warnings,
        "basis": "Годовой план распределён по месяцам равномерно; текущий месяц учитывается пропорционально календарным дням.",
    }


def edit_plan(parent, path, on_saved=None):
    """Small manual editor; no PDF parsing, no spreadsheet dependency."""
    import tkinter as tk
    from tkinter import ttk, messagebox

    path = Path(path)
    if path.exists():
        try:
            initial = load_config(path)
        except AnnualPlanError as exc:
            messagebox.showerror("Годовой план", str(exc), parent=parent)
            return
    else:
        initial = {"version": 1, "asOf": datetime.now(timezone(timedelta(hours=3))).date().isoformat(), "monthPlanTotal": None,
                   "departments": [{"name": name, **dict.fromkeys(FIELDS, 0)} for name in DEFAULT_NAMES]}
    window = tk.Toplevel(parent)
    window.title("Годовой план ДОРЗ • ручной ввод")
    window.geometry("1040x570")
    window.minsize(1000, 570)
    if parent.winfo_viewable():
        window.transient(parent)
    window.configure(padx=24, pady=22)
    ttk.Label(window, text="Годовой план ДОРЗ", font=("", 18)).pack(anchor="w")
    ttk.Label(window, text="Перенесите четыре показателя по каждому подразделению из PDF. Количество — услуги 001983 + 001984.", wraplength=920).pack(anchor="w", pady=(7, 16))
    controls = ttk.Frame(window); controls.pack(fill="x")
    ttk.Label(controls, text="Дата отчёта").pack(side="left")
    date_var = tk.StringVar(value=date.fromisoformat(initial["asOf"]).strftime("%d.%m.%Y"))
    date_entry = ttk.Entry(controls, textvariable=date_var, width=14)
    date_entry.pack(side="left", padx=(10, 24))
    ttk.Label(controls, text="Общий план месяца из PDF").pack(side="left")
    total_var = tk.StringVar(value="" if initial["monthPlanTotal"] is None else str(initial["monthPlanTotal"]))
    ttk.Entry(controls, textvariable=total_var, width=12).pack(side="left", padx=10)
    ttk.Label(window, text="Если общий план месяца оставить пустым, он будет равен сумме подразделений.").pack(anchor="w", pady=(6, 8))
    ttk.Label(window, text="Факт до месяца — с 1 января до начала выбранного месяца. Факт месяца — с 1-го числа по дату отчёта.", wraplength=920).pack(anchor="w", pady=(0, 14))
    holder = ttk.Frame(window); holder.pack(fill="both", expand=True)
    canvas = tk.Canvas(holder, highlightthickness=0)
    scrollbar = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
    table = ttk.Frame(canvas); item = canvas.create_window((0, 0), window=table, anchor="nw")
    canvas.bind("<Configure>", lambda event: canvas.itemconfigure(item, width=event.width))
    table.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))
    for col, label in enumerate(("Подразделение", *LABELS)):
        ttk.Label(table, text=label, wraplength=145 if col else 240).grid(row=0, column=col, sticky="w", padx=4, pady=(0, 8))
        table.columnconfigure(col, weight=2 if col == 0 else 1)
    variables = []
    for number, row in enumerate(initial["departments"], 1):
        values = {key: tk.StringVar(value=str(row[key])) for key in ("name", *FIELDS)}
        variables.append(values)
        for col, key in enumerate(("name", *FIELDS)):
            ttk.Entry(table, textvariable=values[key], width=28 if col == 0 else 12,
                      justify="left" if col == 0 else "right").grid(row=number, column=col, sticky="ew", padx=4, pady=5)
    preview = tk.StringVar(); notes = tk.StringVar()
    ttk.Label(window, textvariable=preview, font=("", 12, "bold"), wraplength=940).pack(anchor="w", pady=(18, 5))
    ttk.Label(window, textvariable=notes, wraplength=940).pack(anchor="w", pady=(0, 14))

    def read_inputs():
        try:
            day = datetime.strptime(date_var.get().strip(), "%d.%m.%Y").date()
        except ValueError:
            raise AnnualPlanError("Дата отчёта должна иметь формат ДД.ММ.ГГГГ.") from None
        return validate_config({"version": 1, "asOf": day.isoformat(), "monthPlanTotal": total_var.get().strip(),
                                "departments": [{key: var.get() for key, var in row.items()} for row in variables]})

    def update_preview(*_):
        try:
            plan = snapshot(read_inputs()); total = plan["total"]
            rate = "план года не задан" if total["yearRate"] is None else f"{total['yearRate']:.1f}% к плану на дату; {total['yearShare']:.1f}% годового плана"
            preview.set((f"Услуг с начала года: {total['yearActual']:,}; " + rate).replace(",", " ").replace(".", ","))
            notes.set(" ".join(plan["warnings"]) or "Проценты и тренд рассчитаны автоматически. Факт с начала года = факт до месяца + факт месяца.")
        except AnnualPlanError as exc:
            preview.set("Проверьте заполнение"); notes.set(str(exc))

    def save():
        try:
            config = save_config(read_inputs(), path)
        except (AnnualPlanError, OSError) as exc:
            messagebox.showerror("Проверьте годовой план", str(exc), parent=window); return
        if on_saved:
            on_saved(snapshot(config))
        window.destroy()

    actions = ttk.Frame(window); actions.pack(fill="x")
    ttk.Button(actions, text="Сохранить годовой план", command=save).pack(side="right")
    ttk.Button(actions, text="Отмена", command=window.destroy).pack(side="right", padx=10)
    for var in (date_var, total_var, *(var for row in variables for var in row.values())):
        var.trace_add("write", update_preview)
    window.bind("<Escape>", lambda event: window.destroy())
    update_preview(); date_entry.focus_set()
    return window
