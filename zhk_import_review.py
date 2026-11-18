"""Review service fields in memory. Never modifies or exports the source workbook."""
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from collections import Counter, defaultdict

from zhk_monthly_report import CENT, clean_text, decimal_number, identify_department, parse_date, service_code

FIELDS = ("date", "code", "name", "quantity", "tariff", "doctor", "department")
LABELS = {"date": "Дата услуги", "code": "Код услуги", "name": "Название услуги",
          "quantity": "Количество", "tariff": "Тариф за 1 услугу, ₽",
          "doctor": "Врач", "department": "Подразделение"}
TARGETS = {"001983", "001984"}


def display(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y")
    return "" if value is None else str(value)


def quantity_number(value):
    quantity = decimal_number(value, "Количество")
    if quantity < 0 or quantity != quantity.to_integral_value() or quantity > 10**9:
        raise ValueError("Нужно целое число от 0 до 1 000 000 000.")
    return quantity


def tariff_number(value):
    price = decimal_number(value, "Тариф")
    if price < 0:
        raise ValueError("Тариф должен быть неотрицательным.")
    return price


def line_amount(quantity, tariff):
    """Use the same per-row rounding as the dashboard collector."""
    return (quantity_number(quantity) * tariff_number(tariff)).quantize(CENT, rounding=ROUND_HALF_UP)


def service_context(values, epoch):
    try:
        code = service_code(values.get("code"))
    except (ValueError, TypeError):
        code = ""
    try:
        day = parse_date(values.get("date"), epoch)
        month = (day.year, day.month)
    except (ValueError, TypeError, OverflowError):
        month = None
    return code, month


@dataclass(frozen=True)
class ServiceSuggestions:
    code: str
    month: tuple | None
    names: tuple
    tariffs: tuple


class ServiceCatalog:
    """Indexed explicit source values; lookup never scans the full workbook.

    Excluded rows are removed, edits update one entry, and the row being edited
    cannot donate its own old name or tariff. Validation-only inferred prices
    are never donors; saved corrections become explicit values in the snapshot.
    """
    def __init__(self, review):
        self.epoch = review.epoch
        self.entries = {}
        self.codes = Counter()
        self.names = defaultdict(Counter)
        self.month_names = defaultdict(Counter)
        self.prices = defaultdict(Counter)
        for number in review.rows:
            if number not in review.excluded:
                self.update(number, review.values(number))

    @staticmethod
    def _adjust(index, key, value, delta):
        counts = index[key]
        counts[value] += delta
        if counts[value] <= 0:
            del counts[value]
        if not counts:
            del index[key]

    def _count(self, entry, delta):
        code, month, name, price = entry
        self.codes[code] += delta
        if self.codes[code] <= 0:
            del self.codes[code]
        if name:
            self._adjust(self.names, code, name, delta)
            if month:
                self._adjust(self.month_names, (code, month), name, delta)
        if price is not None:
            self._adjust(self.prices, (code, month), price, delta)

    def update(self, number, values):
        previous = self.entries.pop(number, None)
        if previous is not None:
            self._count(previous, -1)
        if values is None:
            return
        code, month = service_context(values, self.epoch)
        if not code:
            return
        name, price = clean_text(values.get("name")), None
        if month and clean_text(values.get("department")):
            try:
                quantity_number(values.get("quantity"))
                candidate = tariff_number(values.get("tariff"))
                if candidate > 0:
                    price = candidate
            except (ValueError, TypeError, ArithmeticError):
                pass
        self.entries[number] = entry = code, month, name, price
        self._count(entry, 1)

    @staticmethod
    def _other_values(counts, own):
        return tuple(sorted(value for value, count in counts.items()
                            if count > int(value == own)))

    def lookup(self, number, values):
        code, month = service_context(values, self.epoch)
        own = self.entries.get(number)
        same_code = own is not None and own[0] == code
        same_month = same_code and own[1] == month
        names = self._other_values(self.month_names.get((code, month), {}),
                                   own[2] if same_month else None)
        if not names:
            names = self._other_values(self.names.get(code, {}), own[2] if same_code else None)
        prices = self._other_values(self.prices.get((code, month), {}),
                                    own[3] if same_month else None)
        return ServiceSuggestions(code, month, names, prices)


@dataclass
class PreparedRow:
    number: int
    values: dict
    day: date | None
    selected: bool = False
    department: str = ""
    code: str = ""
    quantity: Decimal | None = None
    price: Decimal | None = None
    restored: bool = False


class ImportReview:
    """A workbook snapshot containing only report fields, never patient columns.

    Validation is the shared gate for both CLI and GUI. All failures are collected,
    including ambiguous missing prices; no row is silently dropped on an error.
    """
    def __init__(self, sheet, epoch, rows, start, end, *, all_departments=False,
                 restore_missing=True, actuals_from_services=False):
        self.sheet, self.epoch = sheet, epoch
        self.rows = {number: values for number, values in rows}
        self.start, self.end = start, end
        self.all_departments = all_departments
        self.restore_missing = restore_missing
        self.actuals_from_services = actuals_from_services
        self.overrides, self.excluded = {}, set()
        self.issues, self.prepared, self.review_rows = {}, [], set()
        self._catalog = None

    def values(self, number):
        return {**self.rows[number], **self.overrides.get(number, {})}

    @property
    def catalog(self):
        if self._catalog is None:
            self._catalog = ServiceCatalog(self)
        return self._catalog

    def suggestions(self, number, values):
        return self.catalog.lookup(number, values)

    def _sync_catalog(self, number):
        if self._catalog is not None:
            self._catalog.update(number, None if number in self.excluded else self.values(number))

    def edit(self, number, values):
        if number not in self.rows or set(values) - set(FIELDS):
            raise ValueError("Неизвестная строка или поле")
        changes = {key: value for key, value in values.items()
                   if display(value) != display(self.rows[number].get(key))}
        if changes:
            self.overrides[number] = changes
        else:
            self.overrides.pop(number, None)
        self.excluded.discard(number)
        self.review_rows.add(number)
        self._sync_catalog(number)

    def exclude(self, numbers):
        for number in numbers:
            if number not in self.rows:
                raise ValueError("Неизвестная строка")
            self.excluded.add(number)
            self.review_rows.add(number)
            self._sync_catalog(number)

    def restore(self, numbers):
        for number in numbers:
            self.excluded.discard(number)
            self.overrides.pop(number, None)
            self._sync_catalog(number)

    def evaluate(self):
        issues, prepared, donors, missing_prices = {}, [], defaultdict(set), []
        first = self.end.replace(day=1)
        for number in self.rows:
            if number in self.excluded:
                continue
            values, errors = self.values(number), {}
            try:
                day = parse_date(values.get("date"), self.epoch)
            except (ValueError, TypeError, OverflowError):
                day = None
                errors["date"] = "Укажите действительную дату, например 17.11.2026."
            raw_department = clean_text(values.get("department"))
            department = identify_department(raw_department, day) if day else None
            in_period = day is None or self.start <= day <= self.end
            in_actuals = self.actuals_from_services and (day is None or first <= day <= self.end)
            selected = in_period and bool(department or self.all_departments)
            # A missing department must not be mistaken for an unrelated department.
            if (in_period or in_actuals) and not raw_department:
                errors["department"] = "Укажите подразделение или исключите строку."
            item = PreparedRow(number, values, day, selected,
                               department or raw_department or "Подразделение не указано")
            relevant = selected or in_actuals or day is None or (in_period and not raw_department)
            if relevant:
                try:
                    item.code = service_code(values.get("code"))
                    if not item.code:
                        raise ValueError("Не заполнен код услуги.")
                except (ValueError, TypeError) as exc:
                    errors["code"] = str(exc)
                if selected or day is None or (in_period and not raw_department) or item.code in TARGETS:
                    try:
                        item.quantity = quantity_number(values.get("quantity"))
                    except (ValueError, TypeError) as exc:
                        errors["quantity"] = str(exc)
                if selected or day is None or (in_period and not raw_department):
                    raw_price = values.get("tariff")
                    if clean_text(raw_price):
                        try:
                            item.price = tariff_number(raw_price)
                        except (ValueError, TypeError) as exc:
                            errors["tariff"] = str(exc)
                    else:
                        missing_prices.append(item)
            if errors:
                issues[number] = errors
            elif selected and item.price is not None and item.price > 0:
                donors[day.year, day.month, item.code].add(item.price)
            prepared.append(item)
        for item in missing_prices:
            if item.quantity == 0:
                item.price, item.restored = Decimal(0), True
                continue
            choices = donors[item.day.year, item.day.month, item.code] if item.day and item.code else set()
            if self.restore_missing and len(choices) == 1:
                item.price, item.restored = next(iter(choices)), True
            else:
                reason = ("В месяце найдены разные цены: " + ", ".join(map(str, sorted(choices)[:5])) + "."
                          if len(choices) > 1 else "Нет единственной цены этого кода в том же месяце.")
                if not self.restore_missing:
                    reason = "Автовосстановление тарифов отключено."
                issues.setdefault(item.number, {})["tariff"] = "Заполните тариф. " + reason
        self.issues, self.prepared = issues, prepared
        self.review_rows.update(issues)
        if issues and self._catalog is None:
            # Initial validation runs in the worker, before the dialog is opened.
            self._catalog = ServiceCatalog(self)
        return not issues

    def audit(self):
        return [{"row": number, "action": "excluded" if number in self.excluded else "edited",
                 "fields": [] if number in self.excluded else sorted(self.overrides[number])}
                for number in sorted(self.excluded | self.overrides.keys())]


class ImportReviewDialog:
    """Native keyboard-friendly review queue; expensive validation is off the Tk loop."""
    def __init__(self, parent, review, finish):
        import tkinter as tk
        from tkinter import ttk
        self.tk, self.ttk, self.review, self.finish = tk, ttk, review, finish
        self.busy, self.current, self.dirty = False, None, set()
        self._loading_fields = False
        self._service_context = ("", None)
        self.window = tk.Toplevel(parent)
        w = self.window
        w.title("Проверка выгрузки • Ритм консультаций")
        width, height = min(1180, w.winfo_screenwidth()-60), min(780, w.winfo_screenheight()-80)
        w.geometry(f"{width}x{height}")
        w.minsize(min(920, width), min(650, height))
        w.transient(parent); w.protocol("WM_DELETE_WINDOW", self.cancel)
        w.configure(background="#F0F4F8")
        style = ttk.Style(w)
        style.configure("Review.Treeview", rowheight=40)
        style.configure("Review.TLabel", background="#F0F4F8")
        style.configure("ReviewTitle.TLabel", background="#F0F4F8", font=("", 19, "bold"))
        outer = ttk.Frame(w, padding=20); outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Проверьте строки выгрузки", font=("", 19, "bold")).pack(anchor="w")
        ttk.Label(outer, text=f"Лист «{review.sheet}». Исправьте поля справа или исключите выбранные строки из отчёта.",
                  wraplength=1000).pack(anchor="w", pady=(6, 4))
        ttk.Label(outer, text="Исходный Excel не изменяется. Полностью пустые строки пропускаются автоматически.",
                  foreground="#50627A", wraplength=1000).pack(anchor="w", pady=(0, 12))
        self.status = tk.StringVar()
        ttk.Label(outer, textvariable=self.status, foreground="#1D5F92", wraplength=1000).pack(anchor="w", pady=(0, 12))
        panes = ttk.Panedwindow(outer, orient="horizontal"); panes.pack(fill="both", expand=True)
        left, right = ttk.Frame(panes), ttk.Frame(panes, padding=(18, 0, 0, 0))
        panes.add(left, weight=3); panes.add(right, weight=2)
        self.tree = ttk.Treeview(left, columns=("row", "state", "problem"), show="headings",
                                 selectmode="extended", style="Review.Treeview", height=10)
        for key, text, width_ in [("row","Строка Excel",95),("state","Состояние",145),("problem","Что проверить",300)]:
            self.tree.heading(key,text=text);self.tree.column(key,width=width_,minwidth=70,stretch=key=="problem")
        self.tree.column("row",anchor="e")
        self.tree.tag_configure("error", foreground="#943D35", background="#FFF4EC")
        self.tree.tag_configure("excluded", foreground="#77818D")
        self.tree.tag_configure("changed", foreground="#1C668F")
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0,column=0,sticky="nsew");scroll.grid(row=0,column=1,sticky="ns")
        left.rowconfigure(0,weight=1);left.columnconfigure(0,weight=1)
        bar = ttk.Frame(left);bar.grid(row=1,column=0,columnspan=2,sticky="ew",pady=(10,0))
        self.exclude_button=ttk.Button(bar,text="Исключить выбранные",command=self.exclude);self.exclude_button.pack(side="left")
        self.restore_button=ttk.Button(bar,text="Вернуть исходные",command=self.restore);self.restore_button.pack(side="left",padx=8)
        ttk.Label(left,text="Ctrl / Shift: выбор нескольких строк. Исключение можно отменить.",
                  foreground="#50627A",wraplength=550).grid(row=2,column=0,columnspan=2,sticky="w",pady=(8,0))
        self.row_title=tk.StringVar(value="Выберите одну строку")
        ttk.Label(right,textvariable=self.row_title,font=("",13,"bold")).pack(anchor="w",pady=(0,8))
        self.fields, self.entries, self.errors = {}, {}, {}
        choices={key:sorted({display(values.get(key)) for values in review.rows.values()
                             if clean_text(values.get(key))}) for key in ("doctor","department")}
        choices.update(code=sorted(review.catalog.codes), name=(), tariff=())
        # Error labels are compact; the full explanation is displayed below the form.
        form=ttk.Frame(right);form.pack(fill="x")
        for i,key in enumerate(FIELDS):
            label=tk.StringVar(value=LABELS[key]);self.errors[key]=label
            ttk.Label(form,textvariable=label).grid(row=i*2,column=0,sticky="w",pady=(3,2))
            value=tk.StringVar();self.fields[key]=value
            entry=(ttk.Combobox(form,textvariable=value,width=37,values=choices[key]) if key in choices
                   else ttk.Entry(form,textvariable=value,width=39))
            entry.grid(row=i*2+1,column=0,sticky="ew")
            entry.bind("<Return>",lambda event:self.save())
            self.entries[key]=entry
        form.columnconfigure(0,weight=1)
        self.amount_text=tk.StringVar(value="Сумма строки: —")
        ttk.Label(right,textvariable=self.amount_text,font=("",12,"bold"),
                  foreground="#175C9C").pack(anchor="w",pady=(10,0))
        self.explanation=tk.StringVar()
        explanation_frame=ttk.Frame(right);explanation_frame.pack(fill="both",expand=True,pady=(8,6))
        explanation_text=tk.Text(explanation_frame,height=4,width=36,wrap="word",font="TkDefaultFont",
                                 relief="flat",background="#FFF4EC",foreground="#943D35",state="disabled")
        self.explanation_text=explanation_text
        explanation_scroll=ttk.Scrollbar(explanation_frame,orient="vertical",command=explanation_text.yview)
        explanation_text.configure(yscrollcommand=explanation_scroll.set)
        explanation_text.pack(side="left",fill="both",expand=True);explanation_scroll.pack(side="right",fill="y")
        def show_explanation(*_):
            explanation_text.configure(state="normal");explanation_text.delete("1.0","end")
            explanation_text.insert("1.0",self.explanation.get());explanation_text.configure(state="disabled")
        self.explanation.trace_add("write",show_explanation)
        self.save_button=ttk.Button(right,text="Применить изменения",command=self.save)
        self.save_button.pack(side="bottom",anchor="w",pady=(6,0),before=form)
        footer=ttk.Frame(outer);footer.pack(side="bottom",fill="x",pady=(16,0),before=panes)
        self.cancel_button=ttk.Button(footer,text="Отменить сборку",command=self.cancel);self.cancel_button.pack(side="left")
        self.continue_button=ttk.Button(footer,text="Проверить и продолжить",command=self.proceed);self.continue_button.pack(side="right")
        self.progress=ttk.Progressbar(footer,mode="indeterminate",length=130)
        self.tree.bind("<<TreeviewSelect>>",self.select)
        self.tree.bind("<Delete>",lambda e:self.exclude())
        self.tree.bind("<Control-a>",self.select_all)
        w.bind("<Control-Return>",lambda e:self.proceed())
        w.bind("<Escape>",lambda e:self.cancel())
        for value in self.fields.values():
            value.trace_add("write", self.field_changed)
        self.entries["code"].bind("<<ComboboxSelected>>", self.autofill_service)
        for key in ("code", "date"):
            self.entries[key].bind("<FocusOut>", self.autofill_service)
        self.refresh();w.grab_set();self.tree.focus_set()

    def form_values(self):
        return {key: value.get() for key, value in self.fields.items()}

    def field_changed(self, *_):
        if not self._loading_fields and not self.busy:
            self.update_form()

    def autofill_service(self, event=None, *, only_changed=False):
        if self._loading_fields or self.busy or self.current is None or self.current in self.review.excluded:
            return
        values = self.form_values()
        context = service_context(values, self.review.epoch)
        code_changed = context[0] != self._service_context[0]
        month_changed = context[1] != self._service_context[1]
        if only_changed and not (code_changed or month_changed):
            return
        choices = self.review.suggestions(self.current, values)
        self._loading_fields = True
        try:
            if code_changed or not clean_text(values["name"]):
                if len(choices.names) == 1:
                    self.fields["name"].set(choices.names[0])
                elif code_changed:
                    self.fields["name"].set("")
            try:
                tariff_number(values["tariff"])
                valid_price = True
            except (ValueError, TypeError, ArithmeticError):
                valid_price = False
            if code_changed or month_changed or not valid_price:
                if self.review.restore_missing and len(choices.tariffs) == 1:
                    self.fields["tariff"].set(format(choices.tariffs[0], "f").replace(".", ","))
                elif code_changed or month_changed:
                    # Never leave the previous service/month's price in the form.
                    self.fields["tariff"].set("")
            self._service_context = context
        finally:
            self._loading_fields = False
        self.update_form()

    def update_form(self):
        active = self.current is not None and self.current not in self.review.excluded
        values = self.form_values()
        choices = self.review.suggestions(self.current, values) if active else None
        self.entries["name"].configure(values=choices.names if choices else ())
        self.entries["tariff"].configure(values=tuple(format(x, "f").replace(".", ",")
                                                     for x in choices.tariffs) if choices else ())
        try:
            amount = line_amount(values["quantity"], values["tariff"]) if active else None
        except (ValueError, TypeError, ArithmeticError):
            amount = None
        amount_label = format(amount, ",.2f").replace(",", " ").replace(".", ",") + " ₽" if amount is not None else "—"
        self.amount_text.set("Сумма строки: " + amount_label)
        current_values = self.review.values(self.current) if active else {}
        pending = active and (self.current in self.dirty or any(
            display(value) != display(current_values.get(key)) for key, value in values.items()))
        errors = self.review.issues.get(self.current, {}) if active and not pending else {}
        for key in FIELDS:
            self.errors[key].set(LABELS[key] + (" · проверьте" if key in errors else ""))
        messages = list(errors.values())
        if self.current in self.review.excluded:
            messages = ["Строка исключена. «Вернуть исходные» отменяет исключение и правки."]
        elif choices and choices.code:
            if len(choices.names) > 1:
                messages.append("У кода несколько названий: выберите нужное из списка.")
            if choices.month is None:
                messages.append("Для подстановки тарифа нужна корректная дата услуги.")
            elif len(choices.tariffs) > 1:
                messages.append("В этом месяце несколько тарифов: выберите нужный из списка.")
            elif not choices.tariffs:
                messages.append("В других строках за этот месяц тариф не найден. Укажите его вручную.")
            elif not self.review.restore_missing:
                messages.append("Автоподстановка тарифа отключена. Его можно выбрать из списка.")
        if pending:
            messages.append("Изменения ожидают проверки: нажмите «Проверить и продолжить».")
        if not messages and active:
            messages.append("Тариф — цена за 1 услугу. Сумма = количество × тариф.")
        self.explanation_text.configure(foreground="#943D35" if errors else "#50627A",
                                        background="#FFF4EC" if errors else "#EDF4FA")
        self.explanation.set("\n".join(messages))

    def numbers(self):
        return [int(x) for x in self.tree.selection()]

    def select_all(self, event=None):
        self.tree.selection_set(self.tree.get_children());return "break"

    def stage(self):
        if self.current is None or self.current in self.review.excluded:
            return
        self.autofill_service(only_changed=True)
        values=self.form_values()
        current=self.review.values(self.current)
        if any(values[k]!=display(current.get(k)) for k in FIELDS):
            self.review.edit(self.current,values);self.dirty.add(self.current)
            self.tree.set(str(self.current),"state","Изменена")
            self.tree.item(str(self.current),tags=("changed",))

    def select(self, event=None):
        if self.busy:return
        self.stage()
        selected=self.numbers();self.current=selected[0] if len(selected)==1 else None
        can_edit=self.current is not None and self.current not in self.review.excluded
        for entry in self.entries.values():entry.configure(state="normal" if can_edit else "disabled")
        self.save_button.configure(state="normal" if can_edit else "disabled")
        self.row_title.set(f"Строка {self.current}" if self.current else f"Выбрано строк: {len(selected)}")
        values=self.review.values(self.current) if self.current else {}
        self._loading_fields = True
        try:
            for key,var in self.fields.items():
                var.set(display(values.get(key)))
            self._service_context = service_context(values, self.review.epoch)
        finally:
            self._loading_fields = False
        self.entries["code"].configure(values=sorted(self.review.catalog.codes))
        self.update_form()
        self.update_status()

    def update_status(self):
        pending=len(set(self.review.issues)-self.review.excluded-self.dirty)
        edited=len(self.review.overrides.keys()-self.review.excluded)
        self.status.set(f"Без решения: {pending}  ·  Изменено: {edited}  ·  Исключено: {len(self.review.excluded)}"
                        + (". Нажмите «Проверить и продолжить», чтобы проверить изменения." if self.dirty else ""))

    def refresh(self):
        selected=self.tree.selection()
        self.current=None
        self.tree.delete(*self.tree.get_children())
        for number in sorted(self.review.review_rows):
            if number in self.review.excluded:state,tag="Исключена","excluded"
            elif number in self.dirty:state,tag="Изменена","changed"
            elif number in self.review.issues:state,tag="Нужна правка","error"
            else:state,tag="Исправлена","changed"
            errors=self.review.issues.get(number,{})
            text=("Нужна повторная проверка" if number in self.dirty else
                  "; ".join(LABELS[k] for k in errors) if errors else "Проверка пройдена")
            self.tree.insert("", "end", iid=str(number), values=(number,state,text),tags=(tag,))
        items=self.tree.get_children();keep=[x for x in selected if x in items]
        if keep:self.tree.selection_set(keep)
        elif items:self.tree.selection_set(items[0])
        self.select();self.update_status()

    def save(self):
        if self.busy:return
        self.stage();self.refresh()

    def exclude(self):
        if self.busy:return
        self.stage();self.review.exclude(self.numbers());self.refresh()

    def restore(self):
        if self.busy:return
        numbers=self.numbers();self.current=None
        self.review.restore(numbers);self.dirty.difference_update(numbers);self.refresh()

    def proceed(self):
        if self.busy:return
        import queue, threading
        self.stage();self.busy=True
        for widget in [*self.entries.values(),self.save_button,self.exclude_button,self.restore_button,self.continue_button]:
            widget.configure(state="disabled")
        self.status.set("Проверяем все исправленные строки и пересчитываем тарифы…")
        self.progress.pack(side="right",padx=12);self.progress.start(12)
        result=queue.Queue()
        def validate():
            try:result.put((self.review.evaluate(),None))
            except Exception as exc:result.put((False,str(exc)))
        threading.Thread(target=validate,daemon=True).start()
        def poll():
            if not self.window.winfo_exists():return
            try:valid,error=result.get_nowait()
            except queue.Empty:self.window.after(60,poll);return
            self.busy=False;self.progress.stop();self.progress.pack_forget();self.dirty.clear()
            if valid:self.close(True);return
            for widget in [self.exclude_button,self.restore_button,self.continue_button]:widget.configure(state="normal")
            self.refresh()
            if error:self.status.set("Проверка не завершена: "+error)
        self.window.after(60,poll)

    def cancel(self):
        from tkinter import messagebox
        if not self.busy:self.stage()
        if self.review.overrides or self.review.excluded or self.dirty:
            if not messagebox.askyesno("Отменить сборку?","Исправления этой сборки будут потеряны. Исходный Excel не изменён.",parent=self.window):return
        self.close(False)

    def close(self, accepted):
        self.window.grab_release();self.window.destroy();self.finish(accepted)
