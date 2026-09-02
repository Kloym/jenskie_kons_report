from __future__ import annotations

import calendar
import io
import os
import queue
import subprocess
import sys
import threading
from contextlib import redirect_stderr, redirect_stdout
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Callable, Sequence, cast

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
except ImportError as exc:
    raise SystemExit(
        "Не найден tkinter. Установите стандартную сборку Python с поддержкой Tcl/Tk."
    ) from exc


APP_TITLE = "Отчёты по услугам женских консультаций"
WINDOW_SIZE = "820x610"
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}

MONTH_NAMES = (
    "Январь",
    "Февраль",
    "Март",
    "Апрель",
    "Май",
    "Июнь",
    "Июль",
    "Август",
    "Сентябрь",
    "Октябрь",
    "Ноябрь",
    "Декабрь",
)

WEEKDAY_NAMES = ("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс")


def parse_user_date(value: str) -> date:
    """Проверяет дату, введённую в поле GUI."""
    value = value.strip()
    for date_format in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, date_format).date()
        except ValueError:
            continue
    raise ValueError("Введите дату в формате ДД.ММ.ГГГГ, например 31.08.2026.")


class CalendarPopup(tk.Toplevel):
    """Небольшой календарь без внешней зависимости tkcalendar."""

    def __init__(
        self,
        parent: tk.Misc,
        initial_date: date,
        on_select: Callable[[date], None],
        anchor_widget: tk.Widget,
    ) -> None:
        super().__init__(parent)
        self.withdraw()
        self.title("Выберите дату")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())

        self._selected_date = initial_date
        self._displayed_month = initial_date.replace(day=1)
        self._on_select = on_select

        self._build_widgets()
        self._render_month()
        self._position_near(anchor_widget)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Left>", lambda _event: self._move_month(-1))
        self.bind("<Right>", lambda _event: self._move_month(1))

        self.deiconify()
        self.lift()
        self.focus_force()
        self.grab_set()

    def _build_widgets(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.grid(row=0, column=0, sticky="nsew")

        navigation = ttk.Frame(container)
        navigation.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        navigation.columnconfigure(1, weight=1)

        ttk.Button(
            navigation,
            text="‹",
            width=3,
            command=lambda: self._move_month(-1),
            style="CalendarNav.TButton",
        ).grid(row=0, column=0)

        self._month_label = ttk.Label(
            navigation,
            anchor="center",
            style="CalendarTitle.TLabel",
        )
        self._month_label.grid(row=0, column=1, sticky="ew", padx=8)

        ttk.Button(
            navigation,
            text="›",
            width=3,
            command=lambda: self._move_month(1),
            style="CalendarNav.TButton",
        ).grid(row=0, column=2)

        weekdays = ttk.Frame(container)
        weekdays.grid(row=1, column=0, sticky="ew")
        for column, name in enumerate(WEEKDAY_NAMES):
            style = "WeekendHeader.TLabel" if column >= 5 else "CalendarHeader.TLabel"
            ttk.Label(
                weekdays,
                text=name,
                width=4,
                anchor="center",
                style=style,
            ).grid(row=0, column=column, padx=1, pady=1)

        self._days_frame = ttk.Frame(container)
        self._days_frame.grid(row=2, column=0, pady=(3, 7))

        ttk.Button(
            container,
            text="Сегодня",
            command=self._select_today,
        ).grid(row=3, column=0, sticky="ew")

    def _render_month(self) -> None:
        for child in self._days_frame.winfo_children():
            child.destroy()

        year = self._displayed_month.year
        month = self._displayed_month.month
        self._month_label.configure(text=f"{MONTH_NAMES[month - 1]} {year}")

        weeks = calendar.Calendar(firstweekday=calendar.MONDAY).monthdayscalendar(
            year, month
        )
        today = date.today()

        for row_index, week in enumerate(weeks):
            for column_index, day_number in enumerate(week):
                if day_number == 0:
                    ttk.Label(self._days_frame, text="", width=4).grid(
                        row=row_index,
                        column=column_index,
                        padx=1,
                        pady=1,
                    )
                    continue

                candidate = date(year, month, day_number)
                if candidate == self._selected_date:
                    style = "SelectedDay.TButton"
                elif candidate == today:
                    style = "Today.TButton"
                elif column_index >= 5:
                    style = "Weekend.TButton"
                else:
                    style = "CalendarDay.TButton"

                ttk.Button(
                    self._days_frame,
                    text=str(day_number),
                    width=4,
                    style=style,
                    command=lambda selected=candidate: self._select_date(selected),
                ).grid(
                    row=row_index,
                    column=column_index,
                    padx=1,
                    pady=1,
                )

    def _move_month(self, delta: int) -> None:
        current_index = self._displayed_month.year * 12 + self._displayed_month.month - 1
        new_index = current_index + delta
        year, zero_based_month = divmod(new_index, 12)
        self._displayed_month = date(year, zero_based_month + 1, 1)
        self._render_month()

    def _select_date(self, selected_date: date) -> None:
        self._on_select(selected_date)
        self.destroy()

    def _select_today(self) -> None:
        self._select_date(date.today())

    def _position_near(self, anchor_widget: tk.Widget) -> None:
        self.update_idletasks()
        x = anchor_widget.winfo_rootx()
        y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height() + 4

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        popup_width = self.winfo_reqwidth()
        popup_height = self.winfo_reqheight()

        x = min(x, screen_width - popup_width - 12)
        y = min(y, screen_height - popup_height - 48)
        self.geometry(f"+{max(8, x)}+{max(8, y)}")


class ExcelReportGUI:
    """Основное окно приложения."""

    def __init__(
        self,
        root: tk.Tk,
        application_factory: Callable[[], object],
    ) -> None:
        self.root = root
        self.application_factory = application_factory
        self._running = False
        self._last_output_dir: Path | None = None
        self._interactive_widgets: list[tk.Widget] = []
        self._result_queue: queue.Queue[tuple[str, object, str]] = queue.Queue()

        today = date.today().strftime("%d.%m.%Y")
        self.source_var = tk.StringVar()
        self.date_from_var = tk.StringVar(value=today)
        self.date_to_var = tk.StringVar(value=today)
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.all_departments_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Готово к работе")

        self._configure_window()
        self._configure_styles()
        self._build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_window(self) -> None:
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(760, 560)
        self.root.configure(background="#F4F7FB")
        self.root.option_add("*tearOff", False)
        self._center_window()

    def _center_window(self) -> None:
        self.root.update_idletasks()
        width = 820
        height = 610
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        available_themes = style.theme_names()
        if "clam" in available_themes:
            style.theme_use("clam")

        style.configure("App.TFrame", background="#F4F7FB")
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure(
            "Title.TLabel",
            background="#F4F7FB",
            foreground="#173B5E",
            font=("Segoe UI", 18, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background="#F4F7FB",
            foreground="#5C7083",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Field.TLabel",
            background="#FFFFFF",
            foreground="#2C4052",
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "Hint.TLabel",
            background="#FFFFFF",
            foreground="#718396",
            font=("Segoe UI", 8),
        )
        style.configure("TEntry", padding=7, font=("Segoe UI", 10))
        style.configure("TButton", padding=(10, 7), font=("Segoe UI", 9))
        style.configure(
            "Accent.TButton",
            background="#207C7E",
            foreground="#FFFFFF",
            padding=(18, 9),
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#176769"), ("disabled", "#A7B9BA")],
            foreground=[("disabled", "#EEF3F3")],
        )
        style.configure(
            "CalendarTitle.TLabel",
            font=("Segoe UI", 10, "bold"),
            foreground="#173B5E",
        )
        style.configure(
            "CalendarHeader.TLabel",
            font=("Segoe UI", 8, "bold"),
            foreground="#53697D",
        )
        style.configure(
            "WeekendHeader.TLabel",
            font=("Segoe UI", 8, "bold"),
            foreground="#A34F68",
        )
        style.configure("CalendarNav.TButton", font=("Segoe UI", 12, "bold"))
        style.configure("CalendarDay.TButton", padding=4)
        style.configure("Weekend.TButton", foreground="#A34F68", padding=4)
        style.configure(
            "Today.TButton",
            foreground="#207C7E",
            font=("Segoe UI", 9, "bold"),
            padding=4,
        )
        style.configure(
            "SelectedDay.TButton",
            background="#207C7E",
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            padding=4,
        )
        style.map(
            "SelectedDay.TButton",
            background=[("active", "#176769")],
            foreground=[("active", "#FFFFFF")],
        )
        style.configure(
            "Status.TLabel",
            background="#E8EEF4",
            foreground="#486074",
            padding=(10, 6),
            font=("Segoe UI", 9),
        )
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#DCE5ED",
            background="#207C7E",
        )

    def _build_layout(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=(22, 18, 22, 14))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)

        ttk.Label(
            outer,
            text="Генератор Excel-отчётов",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            outer,
            text="Выберите файл, задайте период и получите отдельный отчёт для каждой ЖК.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 14))

        card = ttk.Frame(outer, style="Card.TFrame", padding=18)
        card.grid(row=2, column=0, sticky="ew")
        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="Исходный Excel-файл", style="Field.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        source_row = ttk.Frame(card, style="Card.TFrame")
        source_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 2))
        source_row.columnconfigure(0, weight=1)
        source_entry = ttk.Entry(source_row, textvariable=self.source_var)
        source_entry.grid(row=0, column=0, sticky="ew")
        browse_source_button = ttk.Button(
            source_row,
            text="Выбрать файл…",
            command=self._choose_source,
        )
        browse_source_button.grid(row=0, column=1, padx=(8, 0))
        ttk.Label(
            card,
            text="Поддерживаются .xls, .xlsx и .xlsm",
            style="Hint.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky="w")

        ttk.Label(card, text="Дата начала", style="Field.TLabel").grid(
            row=3, column=0, sticky="w", pady=(14, 0)
        )
        ttk.Label(card, text="Дата окончания", style="Field.TLabel").grid(
            row=3, column=1, sticky="w", padx=(12, 0), pady=(14, 0)
        )

        date_from_entry, date_from_button = self._build_date_field(
            card,
            row=4,
            column=0,
            variable=self.date_from_var,
            padding=(0, 0),
        )
        date_to_entry, date_to_button = self._build_date_field(
            card,
            row=4,
            column=1,
            variable=self.date_to_var,
            padding=(12, 0),
        )

        ttk.Label(card, text="Папка результатов", style="Field.TLabel").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(14, 0)
        )
        output_row = ttk.Frame(card, style="Card.TFrame")
        output_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(5, 2))
        output_row.columnconfigure(0, weight=1)
        output_entry = ttk.Entry(output_row, textvariable=self.output_var)
        output_entry.grid(row=0, column=0, sticky="ew")
        browse_output_button = ttk.Button(
            output_row,
            text="Выбрать папку…",
            command=self._choose_output,
        )
        browse_output_button.grid(row=0, column=1, padx=(8, 0))
        ttk.Label(
            card,
            text="Можно оставить пустым — отчёты появятся рядом с исходным файлом.",
            style="Hint.TLabel",
        ).grid(row=7, column=0, columnspan=2, sticky="w")

        options = ttk.Frame(card, style="Card.TFrame")
        options.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        options.columnconfigure(1, weight=1)
        all_departments_check = ttk.Checkbutton(
            options,
            text="Формировать по всем подразделениям",
            variable=self.all_departments_var,
        )
        all_departments_check.grid(row=0, column=0, sticky="w")
        ttk.Label(options, text="Лист:", style="Field.TLabel").grid(
            row=0, column=1, sticky="e", padx=(20, 6)
        )
        sheet_entry = ttk.Entry(options, textvariable=self.sheet_var, width=18)
        sheet_entry.grid(row=0, column=2, sticky="e")
        ttk.Label(
            options,
            text="необязательно",
            style="Hint.TLabel",
        ).grid(row=1, column=2, sticky="e", pady=(2, 0))

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        actions.columnconfigure(0, weight=1)
        self.open_output_button = ttk.Button(
            actions,
            text="Открыть папку",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_output_button.grid(row=0, column=1, padx=(0, 8))
        self.start_button = ttk.Button(
            actions,
            text="Сформировать отчёты",
            command=self._start_processing,
            style="Accent.TButton",
        )
        self.start_button.grid(row=0, column=2)

        log_frame = ttk.Frame(outer, style="App.TFrame")
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(14, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        ttk.Label(
            log_frame,
            text="Журнал обработки",
            style="Subtitle.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            wrap="word",
            font=("Consolas", 9),
            background="#FFFFFF",
            foreground="#263A43",
            relief="flat",
            borderwidth=1,
        )
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

        bottom = ttk.Frame(outer, style="App.TFrame")
        bottom.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        bottom.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=180)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Label(bottom, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0, column=1, sticky="e"
        )

        self._interactive_widgets.extend(
            [
                source_entry,
                browse_source_button,
                date_from_entry,
                date_from_button,
                date_to_entry,
                date_to_button,
                output_entry,
                browse_output_button,
                all_departments_check,
                sheet_entry,
                self.start_button,
            ]
        )

    def _build_date_field(
        self,
        parent: ttk.Frame,
        row: int,
        column: int,
        variable: tk.StringVar,
        padding: tuple[int, int],
    ) -> tuple[ttk.Entry, ttk.Button]:
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.grid(
            row=row,
            column=column,
            sticky="ew",
            padx=padding,
            pady=(5, 0),
        )
        frame.columnconfigure(0, weight=1)
        entry = ttk.Entry(frame, textvariable=variable)
        entry.grid(row=0, column=0, sticky="ew")
        button = ttk.Button(frame, text="📅", width=4)
        button.configure(
            command=lambda: self._open_calendar(variable, button)
        )
        button.grid(row=0, column=1, padx=(6, 0))
        return entry, button

    def _open_calendar(self, variable: tk.StringVar, anchor: tk.Widget) -> None:
        try:
            initial_date = parse_user_date(variable.get())
        except ValueError:
            initial_date = date.today()

        CalendarPopup(
            self.root,
            initial_date,
            lambda selected: variable.set(selected.strftime("%d.%m.%Y")),
            anchor,
        )

    def _choose_source(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="Выберите таблицу с услугами",
            filetypes=(
                ("Excel-файлы", "*.xls *.xlsx *.xlsm"),
                ("Все файлы", "*.*"),
            ),
        )
        if filename:
            self.source_var.set(filename)

    def _choose_output(self) -> None:
        initial_dir = self.output_var.get().strip()
        if not initial_dir and self.source_var.get().strip():
            initial_dir = str(Path(self.source_var.get().strip()).parent)
        directory = filedialog.askdirectory(
            parent=self.root,
            title="Выберите папку для результатов",
            initialdir=initial_dir or None,
        )
        if directory:
            self.output_var.set(directory)

    def _validate_form(self) -> tuple[Path, date, date, Path | None]:
        raw_source = self.source_var.get().strip().strip('"')
        if not raw_source:
            raise ValueError("Выберите исходный Excel-файл.")

        source_path = Path(raw_source).expanduser()
        if not source_path.exists() or not source_path.is_file():
            raise ValueError(f"Файл не найден:\n{source_path}")
        if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError("Поддерживаются файлы .xls, .xlsx и .xlsm.")

        date_from = parse_user_date(self.date_from_var.get())
        date_to = parse_user_date(self.date_to_var.get())
        if date_from > date_to:
            raise ValueError("Дата начала не может быть позже даты окончания.")

        raw_output = self.output_var.get().strip().strip('"')
        output_dir = Path(raw_output).expanduser() if raw_output else None
        if output_dir is not None and output_dir.exists() and not output_dir.is_dir():
            raise ValueError("Путь для результатов должен указывать на папку.")

        return source_path, date_from, date_to, output_dir

    def _start_processing(self) -> None:
        if self._running:
            return

        try:
            source_path, date_from, date_to, output_dir = self._validate_form()
        except ValueError as exc:
            messagebox.showerror("Проверьте данные", str(exc), parent=self.root)
            return

        arguments = SimpleNamespace(
            input=str(source_path),
            date_from=date_from.strftime("%d.%m.%Y"),
            date_to=date_to.strftime("%d.%m.%Y"),
            output_dir=str(output_dir) if output_dir is not None else None,
            sheet=self.sheet_var.get().strip() or None,
            all_departments=bool(self.all_departments_var.get()),
            no_pause=True,
        )

        self._running = True
        self._last_output_dir = None
        self.open_output_button.state(["disabled"])
        self._set_form_enabled(False)
        self._clear_log()
        self._append_log("Запуск обработки…\n")
        self.status_var.set("Обработка данных…")
        self.progress.start(12)

        worker = threading.Thread(
            target=self._run_worker,
            args=(arguments,),
            daemon=True,
        )
        worker.start()
        self.root.after(100, self._poll_worker_result)

    def _run_worker(self, arguments: SimpleNamespace) -> None:
        console_output = io.StringIO()
        try:
            application = self.application_factory()
            with redirect_stdout(console_output), redirect_stderr(console_output):
                files = application.run(arguments)
        except Exception as exc:
            self._result_queue.put(
                ("error", str(exc), console_output.getvalue())
            )
            return

        self._result_queue.put(
            ("success", files, console_output.getvalue())
        )

    def _poll_worker_result(self) -> None:
        try:
            result_type, payload, console_output = self._result_queue.get_nowait()
        except queue.Empty:
            if self._running:
                self.root.after(100, self._poll_worker_result)
            return

        if result_type == "success":
            self._finish_successfully(
                cast(Sequence[Path], payload),
                console_output,
            )
        else:
            self._finish_with_error(str(payload), console_output)

    def _finish_successfully(
        self,
        files: Sequence[Path],
        console_output: str,
    ) -> None:
        self._finish_common()
        if console_output.strip():
            self._append_log(console_output.strip() + "\n")

        if files:
            self._last_output_dir = Path(files[0]).parent
            self.open_output_button.state(["!disabled"])
        self.status_var.set(f"Готово: создано файлов — {len(files)}")
        messagebox.showinfo(
            "Отчёты созданы",
            f"Обработка завершена.\nСоздано файлов: {len(files)}",
            parent=self.root,
        )

    def _finish_with_error(self, error_message: str, console_output: str) -> None:
        self._finish_common()
        if console_output.strip():
            self._append_log(console_output.strip() + "\n")
        self._append_log(f"Ошибка: {error_message}\n")
        self.status_var.set("Ошибка обработки")
        messagebox.showerror(
            "Не удалось создать отчёты",
            error_message or "Произошла неизвестная ошибка.",
            parent=self.root,
        )

    def _finish_common(self) -> None:
        self._running = False
        self.progress.stop()
        self._set_form_enabled(True)

    def _set_form_enabled(self, enabled: bool) -> None:
        state = "!disabled" if enabled else "disabled"
        for widget in self._interactive_widgets:
            try:
                widget.state([state])
            except (AttributeError, tk.TclError):
                widget.configure(state="normal" if enabled else "disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_output_folder(self) -> None:
        if self._last_output_dir is None or not self._last_output_dir.exists():
            messagebox.showwarning(
                "Папка не найдена",
                "Сначала сформируйте отчёты.",
                parent=self.root,
            )
            return

        try:
            if sys.platform == "win32":
                os.startfile(self._last_output_dir)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self._last_output_dir)])
            else:
                subprocess.Popen(["xdg-open", str(self._last_output_dir)])
        except OSError as exc:
            messagebox.showerror(
                "Не удалось открыть папку",
                str(exc),
                parent=self.root,
            )

    def _on_close(self) -> None:
        if self._running:
            messagebox.showinfo(
                "Обработка выполняется",
                "Дождитесь окончания создания файлов, затем закройте программу.",
                parent=self.root,
            )
            return
        self.root.destroy()


def launch_gui(
    application_factory: Callable[[], object] | None = None,
) -> None:
    """Запускает GUI отдельно или из основного модуля."""
    if application_factory is None:
        try:
            from zhk_report import ReportApplication
        except ImportError as exc:
            raise SystemExit(
                "Файл zhk_excel_report.py должен находиться рядом с zhk_excel_gui.py."
            ) from exc
        application_factory = ReportApplication

    root = tk.Tk()
    ExcelReportGUI(root, application_factory)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()