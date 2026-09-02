from __future__ import annotations

import argparse
import math
import os
import re
import sys
import xlrd
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Callable, Mapping, Sequence

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils.datetime import from_excel
    from openpyxl.workbook.properties import CalcProperties
except ImportError as exc:
    raise SystemExit(
        "Не найдена библиотека openpyxl. Установите её командой:\n"
        "python -m pip install openpyxl"
    ) from exc


APP_NAME = "Отчёты по услугам женских консультаций"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
DATE_FORMATS = (
    "%d.%m.%Y",
    "%d.%m.%y",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d-%m-%Y",
)
MONEY_QUANT = Decimal("0.01")


class ReportError(Exception):
    """Понятная пользователю ошибка обработки."""


@dataclass(frozen=True)
class AppConfig:
    """Настройки, которые чаще всего требуется менять при сопровождении."""

    preferred_sheet_name: str = "услуги"
    header_scan_rows: int = 30
    header_scan_columns: int = 200
    max_warning_examples: int = 10

    department_predicate: Callable[[str], bool] = field(
        default=lambda value: is_women_consultation(value),
        compare=False,
        repr=False,
    )


@dataclass(frozen=True)
class ColumnMap:
    service_date: int
    service_code: int
    service_name: int
    quantity: int
    tariff: int
    doctor: int
    department: int


@dataclass(frozen=True)
class HeaderInfo:
    row_number: int
    columns: ColumnMap


@dataclass
class ProcessingStats:
    source_rows: int = 0
    rows_in_period: int = 0
    rows_in_selected_departments: int = 0
    aggregated_rows: int = 0
    skipped_rows: int = 0
    blank_quantity_rows: int = 0
    warning_examples: list[str] = field(default_factory=list)

    def warn(self, row_number: int, message: str, limit: int) -> None:
        self.skipped_rows += 1
        if len(self.warning_examples) < limit:
            self.warning_examples.append(f"Строка {row_number}: {message}")


@dataclass
class AggregatedService:
    doctor: str
    doctor_key: str
    service_code: str
    service_name: str
    service_name_key: str
    tariff: Decimal
    quantity: Decimal = Decimal("0")


@dataclass
class DepartmentData:
    name: str
    records: dict[tuple[str, str, str, Decimal], AggregatedService] = field(
        default_factory=dict
    )

    def add(
        self,
        doctor: str,
        service_code: str,
        service_name: str,
        tariff: Decimal,
        quantity: Decimal,
    ) -> None:
        doctor_key = normalize_key(doctor)
        service_name_key = normalize_key(service_name)
        key = (doctor_key, service_code.casefold(), service_name_key, tariff)
        record = self.records.get(key)
        if record is None:
            record = AggregatedService(
                doctor=doctor,
                doctor_key=doctor_key,
                service_code=service_code,
                service_name=service_name,
                service_name_key=service_name_key,
                tariff=tariff,
            )
            self.records[key] = record
        record.quantity += quantity

    def sorted_records(self) -> list[AggregatedService]:
        return sorted(
            self.records.values(),
            key=lambda item: (
                natural_sort_key(item.doctor),
                natural_sort_key(item.service_code),
                natural_sort_key(item.service_name),
                item.tariff,
            ),
        )


def normalize_key(value: object) -> str:
    """Нормализует заголовки и ключи без потери исходного отображения."""
    text = str(value or "").replace("ё", "е").replace("Ё", "Е").casefold()
    return re.sub(r"[^0-9a-zа-я]+", " ", text).strip()


def clean_text(value: object) -> str:
    """Убирает лишние пробелы и недопустимые для Excel управляющие символы."""
    if value is None:
        return ""
    text = ILLEGAL_CHARACTERS_RE.sub("", str(value))
    return re.sub(r"\s+", " ", text).strip()


def safe_excel_text(value: object) -> str:
    """Защищает отчёт от формул, случайно попавших в текстовые поля."""
    text = clean_text(value)
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def is_women_consultation(value: str) -> bool:
    normalized = normalize_key(value)
    by_words = "женск" in normalized and "консультац" in normalized
    by_abbreviation = bool(re.search(r"(?:^|\s)жк(?:\s|\d|$)", normalized))
    return by_words or by_abbreviation


def natural_sort_key(value: object) -> tuple[tuple[int, object], ...]:
    parts = re.split(r"(\d+)", clean_text(value).casefold())
    return tuple((0, int(part)) if part.isdigit() else (1, part) for part in parts)


def parse_date(value: object, *, epoch: datetime | None = None) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            converted = from_excel(value, epoch=epoch)
            return converted.date() if isinstance(converted, datetime) else converted
        except (TypeError, ValueError, OverflowError) as exc:
            raise ValueError(f"некорректная дата {value!r}") from exc

    text = clean_text(value)
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"не удалось распознать дату {text!r}")


def to_decimal(value: object) -> Decimal:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool) or value is None:
        raise ValueError("пустое или логическое значение")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            raise ValueError("нечисловое значение")
        return Decimal(str(value))

    text = clean_text(value).replace("\u00a0", "").replace(" ", "")
    text = re.sub(r"[^0-9,\.\-+]", "", text)
    if not text:
        raise ValueError("пустое значение")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"не удалось распознать число {value!r}") from exc


def excel_number(value: Decimal) -> int | float:
    integral = value.to_integral_value()
    return int(integral) if value == integral else float(value)


def service_code_from_cell(cell: object) -> str:
    value = getattr(cell, "value", None)
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)

    if isinstance(value, (int, float, Decimal)):
        number = Decimal(str(value))
        if number == number.to_integral_value():
            code = str(int(number))
            number_format = str(getattr(cell, "number_format", "") or "")
            primary_format = number_format.split(";")[0].strip()
            if re.fullmatch(r"0+", primary_format):
                code = code.zfill(len(primary_format))
            return code
        return format(number.normalize(), "f")

    text = clean_text(value)
    return text[1:] if text.startswith("'") else text


class HeaderResolver:
    """Находит нужные столбцы даже при небольших различиях в заголовках."""

    ALIASES: Mapping[str, set[str]] = {
        "service_date": {
            "дата",
            "дата услуги",
            "дата оказания услуги",
            "дата выполнения",
        },
        "service_code": {"код услуги", "код мед услуги", "код медуслуги"},
        "service_name": {
            "наименование услуги",
            "название услуги",
            "услуга",
        },
        "quantity": {"кол во", "количество", "количество услуг"},
        "tariff": {
            "тариф",
            "тариф руб",
            "цена",
            "цена руб",
            "цена тариф руб",
        },
        "doctor": {"врач", "фио врача", "врач фио"},
        "department": {
            "подразделение",
            "наименование подразделения",
            "отделение",
        },
    }

    RUSSIAN_NAMES: Mapping[str, str] = {
        "service_date": "Дата",
        "service_code": "Код услуги",
        "service_name": "Наименование услуги",
        "quantity": "Кол-во",
        "tariff": "Тариф",
        "doctor": "Врач",
        "department": "Подразделение",
    }

    def __init__(self, config: AppConfig) -> None:
        self.config = config

    def find(self, worksheet: object) -> HeaderInfo:
        max_col = min(
            int(getattr(worksheet, "max_column", 0) or self.config.header_scan_columns),
            self.config.header_scan_columns,
        )
        best_row = 0
        best_mapping: dict[str, int] = {}

        for row in worksheet.iter_rows(
            min_row=1,
            max_row=self.config.header_scan_rows,
            max_col=max_col,
        ):
            mapping = self._resolve_row([getattr(cell, "value", None) for cell in row])
            if len(mapping) > len(best_mapping):
                best_row = int(getattr(row[0], "row", 0) or 0)
                best_mapping = mapping
            if len(mapping) == len(self.ALIASES):
                return HeaderInfo(
                    row_number=int(getattr(row[0], "row", 0)),
                    columns=ColumnMap(**mapping),
                )

        missing = [
            self.RUSSIAN_NAMES[key]
            for key in self.ALIASES
            if key not in best_mapping
        ]
        found = ", ".join(self.RUSSIAN_NAMES[key] for key in best_mapping) or "нет"
        raise ReportError(
            f"Не найдены обязательные столбцы: {', '.join(missing)}. "
            f"Лучшее совпадение — строка {best_row or 'не определена'}, найдено: {found}."
        )

    def _resolve_row(self, values: Sequence[object]) -> dict[str, int]:
        normalized = [normalize_key(value) for value in values]
        mapping: dict[str, int] = {}
        for field_name, aliases in self.ALIASES.items():
            for column_index, header in enumerate(normalized):
                if header in aliases:
                    mapping[field_name] = column_index
                    break
        return mapping

class XlsCellAdapter:
    """Приводит ячейку старого .xls к интерфейсу openpyxl."""

    def __init__(self, book, sheet, row_index: int, column_index: int) -> None:
        source_cell = sheet.cell(row_index, column_index)

        self.row = row_index + 1
        self.value = source_cell.value
        self.number_format = "General"

        if source_cell.ctype == xlrd.XL_CELL_DATE:
            self.value = xlrd.xldate.xldate_as_datetime(
                source_cell.value,
                book.datemode,
            )
        elif source_cell.ctype == xlrd.XL_CELL_BOOLEAN:
            self.value = bool(source_cell.value)
        elif source_cell.ctype == xlrd.XL_CELL_ERROR:
            self.value = xlrd.error_text_from_code.get(
                source_cell.value,
                f"#ERROR_{source_cell.value}",
            )

        if source_cell.xf_index is not None:
            try:
                xf = book.xf_list[source_cell.xf_index]
                excel_format = book.format_map.get(xf.format_key)
                if excel_format is not None:
                    self.number_format = excel_format.format_str
            except (IndexError, KeyError, AttributeError):
                pass


class XlsWorksheetAdapter:
    """Предоставляет старый лист .xls в формате, понятном SourceReader."""

    def __init__(self, book, sheet) -> None:
        self._book = book
        self._sheet = sheet
        self.title = sheet.name
        self.max_column = sheet.ncols

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        max_col: int | None = None,
    ):
        start_row = max(min_row - 1, 0)
        end_row = min(
            max_row if max_row is not None else self._sheet.nrows,
            self._sheet.nrows,
        )
        end_column = min(
            max_col if max_col is not None else self._sheet.ncols,
            self._sheet.ncols,
        )

        for row_index in range(start_row, end_row):
            yield tuple(
                XlsCellAdapter(
                    self._book,
                    self._sheet,
                    row_index,
                    column_index,
                )
                for column_index in range(end_column)
            )


class XlsWorkbookAdapter:
    """Класс для работы с .xls файлами"""

    def __init__(self, path: Path) -> None:
        self._book = xlrd.open_workbook(
            filename=str(path),
            formatting_info=True,
            on_demand=True,
        )

        self.epoch = (
            datetime(1904, 1, 1)
            if self._book.datemode == 1
            else datetime(1899, 12, 30)
        )

        self.worksheets = [
            XlsWorksheetAdapter(self._book, sheet)
            for sheet in self._book.sheets()
        ]

    def close(self) -> None:
        release_resources = getattr(self._book, "release_resources", None)
        if release_resources is not None:
            release_resources()

class SourceReader:
    """Читает исходную книгу и агрегирует строки в памяти."""

    def __init__(self, config: AppConfig) -> None:
        self.config = config
        self.header_resolver = HeaderResolver(config)

    def aggregate(
        self,
        source_path: Path,
        date_from: date,
        date_to: date,
        *,
        sheet_name: str | None = None,
        include_all_departments: bool = False,
    ) -> tuple[list[DepartmentData], ProcessingStats, str]:
        stats = ProcessingStats()
        try:
            if source_path.suffix.lower() == ".xls":
                if xlrd is None:
                    raise ReportError(
                        "Для обработки .xls установите библиотеку: "
                        "python -m pip install xlrd"
                    )
                workbook = XlsWorkbookAdapter(source_path)
            else:
                workbook = load_workbook(
                    filename=source_path,
                    read_only=True,
                    data_only=True,
                    keep_links=False,
                )
        except ReportError:
            raise
        except Exception as exc:
            raise ReportError(f"Не удалось открыть Excel-файл: {exc}") from exc

        try:
            worksheet, header = self._select_worksheet(workbook, sheet_name)
            departments: dict[str, DepartmentData] = {}
            epoch = getattr(workbook, "epoch", None)

            for row_number, row in enumerate(
                worksheet.iter_rows(min_row=header.row_number + 1),
                start=header.row_number + 1,
            ):
                stats.source_rows += 1
                columns = header.columns

                try:
                    row_date = parse_date(row[columns.service_date].value, epoch=epoch)
                except (ValueError, TypeError, IndexError) as exc:
                    stats.warn(row_number, str(exc), self.config.max_warning_examples)
                    continue

                if not date_from <= row_date <= date_to:
                    continue
                stats.rows_in_period += 1

                try:
                    department = clean_text(row[columns.department].value)
                except IndexError:
                    department = ""
                if not department:
                    stats.warn(row_number, "не заполнено подразделение", self.config.max_warning_examples)
                    continue
                if not include_all_departments and not self.config.department_predicate(department):
                    continue
                stats.rows_in_selected_departments += 1

                try:
                    doctor = clean_text(row[columns.doctor].value)
                    service_name = clean_text(row[columns.service_name].value)
                    service_code = service_code_from_cell(row[columns.service_code])
                    tariff = to_decimal(row[columns.tariff].value).quantize(
                        MONEY_QUANT, rounding=ROUND_HALF_UP
                    )
                    quantity_value = row[columns.quantity].value
                    if quantity_value is None or clean_text(quantity_value) == "":
                        quantity = Decimal("1")
                        stats.blank_quantity_rows += 1
                    else:
                        quantity = to_decimal(quantity_value)
                except (ValueError, TypeError, IndexError) as exc:
                    stats.warn(row_number, f"ошибка в данных услуги: {exc}", self.config.max_warning_examples)
                    continue

                if not doctor:
                    stats.warn(row_number, "не заполнен врач", self.config.max_warning_examples)
                    continue
                if not service_code:
                    stats.warn(row_number, "не заполнен код услуги", self.config.max_warning_examples)
                    continue
                if not service_name:
                    stats.warn(row_number, "не заполнено наименование услуги", self.config.max_warning_examples)
                    continue

                department_key = normalize_key(department)
                bucket = departments.get(department_key)
                if bucket is None:
                    bucket = DepartmentData(name=department)
                    departments[department_key] = bucket
                bucket.add(doctor, service_code, service_name, tariff, quantity)

            stats.aggregated_rows = sum(len(item.records) for item in departments.values())
            return (
                sorted(departments.values(), key=lambda item: natural_sort_key(item.name)),
                stats,
                worksheet.title,
            )
        finally:
            workbook.close()

    def _select_worksheet(
        self, workbook: object, requested_name: str | None
    ) -> tuple[object, HeaderInfo]:
        worksheets = list(workbook.worksheets)
        if requested_name:
            exact = [ws for ws in worksheets if normalize_key(ws.title) == normalize_key(requested_name)]
            if not exact:
                raise ReportError(f"Лист {requested_name!r} не найден в книге.")
            return exact[0], self.header_resolver.find(exact[0])

        preferred_key = normalize_key(self.config.preferred_sheet_name)
        ordered = sorted(
            worksheets,
            key=lambda ws: normalize_key(ws.title) != preferred_key,
        )
        errors: list[str] = []
        for worksheet in ordered:
            try:
                return worksheet, self.header_resolver.find(worksheet)
            except ReportError as exc:
                errors.append(f"{worksheet.title}: {exc}")

        detail = "\n".join(errors[:3])
        raise ReportError(
            "Ни на одном листе не найдена подходящая таблица.\n" + detail
        )


class OutputDirectoryFactory:
    @staticmethod
    def create(parent: Path, date_from: date, date_to: date) -> Path:
        parent.mkdir(parents=True, exist_ok=True)
        base_name = f"Отчёты_ЖК_{date_from:%d.%m.%Y}-{date_to:%d.%m.%Y}"
        candidate = parent / base_name
        if not candidate.exists():
            candidate.mkdir()
            return candidate

        timestamp = datetime.now().strftime("%H%M%S")
        candidate = parent / f"{base_name}_{timestamp}"
        counter = 2
        while candidate.exists():
            candidate = parent / f"{base_name}_{timestamp}_{counter}"
            counter += 1
        candidate.mkdir()
        return candidate


class ReportWorkbookWriter:
    """Создаёт единообразные, печатные и удобные для фильтрации отчёты."""

    NAVY = "173B5E"
    TEAL = "207C7E"
    PALE_BLUE = "DCEAF7"
    PALE_TEAL = "EAF5F3"
    WHITE = "FFFFFF"
    TEXT = "203040"
    MUTED = "587086"
    GRID = "C9D6E2"
    GOLD = "D5A544"

    THEMES = (
        ("173B5E", "207C7E", "DCEAF7", "EAF5F3"),  # Синий / бирюзовый
        ("3E2F5B", "765C9D", "E9E2F3", "F4EFF9"),  # Фиолетовый
        ("214E3B", "3E7C59", "DDEFE5", "EEF7F1"),  # Изумрудный
        ("64283A", "A34F68", "F2DEE5", "F9EDF1"),  # Бордовый
        ("253B6E", "8B5E17", "E2E8F5", "F8F1E4"),  # Индиго / золотой
        ("6A3528", "A94F36", "F2E1DB", "F9EFEA"),  # Терракотовый
        ("263A43", "3E7480", "DFE9EC", "EDF5F6"),  # Графитовый
    )

    def _activate_theme(self, theme_index: int) -> None:
        (
            self.NAVY,
            self.TEAL,
            self.PALE_BLUE,
            self.PALE_TEAL,
        ) = self.THEMES[theme_index % len(self.THEMES)]

    def write_all(
        self,
        departments: Sequence[DepartmentData],
        output_dir: Path,
        source_path: Path,
        source_sheet: str,
        date_from: date,
        date_to: date,
    ) -> list[Path]:
        result: list[Path] = []
        for theme_index, department in enumerate(departments):
            self._activate_theme(theme_index)

            safe_name = sanitize_filename(
                department.name,
                max_length=125,
            )
            filename = (
                f"{safe_name}__{date_from:%d.%m.%Y}-{date_to:%d.%m.%Y}.xlsx"
            )
            destination = output_dir / filename
            self._write_one(
                department,
                destination,
                source_path,
                source_sheet,
                date_from,
                date_to,
            )
            result.append(destination)
        return result

    def _write_one(
        self,
        department: DepartmentData,
        destination: Path,
        source_path: Path,
        source_sheet: str,
        date_from: date,
        date_to: date,
    ) -> None:
        records = department.sorted_records()
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Отчёт"
        worksheet.sheet_view.showGridLines = False
        worksheet.freeze_panes = "A6"
        worksheet.sheet_properties.tabColor = self.TEAL

        workbook.properties.title = department.name
        workbook.properties.subject = (
            f"Сводный отчёт за {date_from:%d.%m.%Y}–{date_to:%d.%m.%Y}"
        )
        workbook.properties.creator = None
        workbook.properties.description = None
        workbook.calculation = CalcProperties(
            calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True
        )

        self._write_title(worksheet, department.name, date_from, date_to)
        first_data_row = 6
        last_data_row = first_data_row + len(records) - 1
        self._write_summary(worksheet, records, first_data_row, last_data_row)
        self._write_headers(worksheet)
        self._write_rows(worksheet, records, first_data_row)
        self._format_layout(worksheet, last_data_row)

        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_suffix(".tmp.xlsx")
        try:
            workbook.save(temporary)
            os.replace(temporary, destination)
        finally:
            workbook.close()
            if temporary.exists():
                temporary.unlink()

    def _write_title(
        self, worksheet: object, department: str, date_from: date, date_to: date
    ) -> None:
        worksheet.merge_cells("A1:F1")
        worksheet["A1"] = "СВОДНЫЙ ОТЧЁТ ПО УСЛУГАМ"
        worksheet["A1"].font = Font(name="Aptos Display", size=16, bold=True, color=self.WHITE)
        worksheet["A1"].fill = PatternFill("solid", fgColor=self.NAVY)
        worksheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
        worksheet.row_dimensions[1].height = 32

        worksheet.merge_cells("A2:F2")
        worksheet["A2"] = (
            f"{safe_excel_text(department)}  •  "
            f"период {date_from:%d.%m.%Y} — {date_to:%d.%m.%Y}"
        )
        worksheet["A2"].font = Font(name="Aptos", size=11, bold=True, color=self.WHITE)
        worksheet["A2"].fill = PatternFill("solid", fgColor=self.TEAL)
        worksheet["A2"].alignment = Alignment(horizontal="center", vertical="center")
        worksheet.row_dimensions[2].height = 26

    def _write_summary(
        self,
        worksheet: object,
        records: Sequence[AggregatedService],
        first_data_row: int,
        last_data_row: int,
    ) -> None:
        doctor_count = len({record.doctor_key for record in records})
        values = {
            "A3": "Врачей",
            "B3": doctor_count,
            "C3": "Всего услуг",
            "D3": f"=SUM(E{first_data_row}:E{last_data_row})",
            "E3": "Итоговая сумма",
            "F3": f"=SUM(F{first_data_row}:F{last_data_row})",
        }
        for coordinate, value in values.items():
            worksheet[coordinate] = value

        label_fill = PatternFill("solid", fgColor=self.PALE_BLUE)
        value_fill = PatternFill("solid", fgColor=self.WHITE)
        for coordinate in ("A3", "C3", "E3"):
            cell = worksheet[coordinate]
            cell.fill = label_fill
            cell.font = Font(name="Aptos", size=10, bold=True, color=self.MUTED)
            cell.alignment = Alignment(horizontal="left", vertical="center")
        for coordinate in ("B3", "D3", "F3"):
            cell = worksheet[coordinate]
            cell.fill = value_fill
            cell.font = Font(name="Aptos Display", size=12, bold=True, color=self.NAVY)
            cell.alignment = Alignment(horizontal="right", vertical="center")

        worksheet["B3"].number_format = "#,##0"
        worksheet["D3"].number_format = "#,##0"
        worksheet["F3"].number_format = '#,##0.00" ₽"'
        worksheet.row_dimensions[3].height = 28
        worksheet.row_dimensions[4].height = 8

    def _write_headers(self, worksheet: object) -> None:
        headers = (
            "Врач",
            "Код услуги",
            "Наименование услуги",
            "Тариф, руб.",
            "Кол-во",
            "Сумма, руб.",
        )
        for column, value in enumerate(headers, start=1):
            cell = worksheet.cell(row=5, column=column, value=value)
            cell.fill = PatternFill("solid", fgColor=self.NAVY)
            cell.font = Font(name="Aptos", size=10, bold=True, color=self.WHITE)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        worksheet.row_dimensions[5].height = 34

    def _write_rows(
        self,
        worksheet: object,
        records: Sequence[AggregatedService],
        first_data_row: int,
    ) -> None:
        thin = Side(style="thin", color=self.GRID)
        group_top = Side(style="medium", color=self.TEAL)
        previous_doctor: str | None = None
        group_index = -1

        for offset, record in enumerate(records):
            row_number = first_data_row + offset
            new_doctor = record.doctor_key != previous_doctor
            if new_doctor:
                group_index += 1
                previous_doctor = record.doctor_key
            fill_color = self.WHITE if group_index % 2 == 0 else self.PALE_TEAL
            fill = PatternFill("solid", fgColor=fill_color)
            top = group_top if new_doctor else thin
            border = Border(top=top, bottom=thin)

            values: Sequence[object] = (
                safe_excel_text(record.doctor),
                safe_excel_text(record.service_code),
                safe_excel_text(record.service_name),
                float(record.tariff),
                excel_number(record.quantity),
                f"=ROUND(D{row_number}*E{row_number},2)",
            )
            for column, value in enumerate(values, start=1):
                cell = worksheet.cell(row=row_number, column=column, value=value)
                cell.fill = fill
                cell.border = border
                cell.font = Font(name="Aptos", size=10, color=self.TEXT)
                cell.alignment = Alignment(vertical="center")

            worksheet.cell(row_number, 1).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
            worksheet.cell(row_number, 2).alignment = Alignment(
                horizontal="center", vertical="center"
            )
            worksheet.cell(row_number, 2).number_format = "@"
            worksheet.cell(row_number, 3).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
            worksheet.cell(row_number, 4).alignment = Alignment(
                horizontal="right", vertical="center"
            )
            worksheet.cell(row_number, 4).number_format = "#,##0.00"
            worksheet.cell(row_number, 5).alignment = Alignment(
                horizontal="right", vertical="center"
            )
            worksheet.cell(row_number, 5).number_format = "#,##0"
            worksheet.cell(row_number, 6).alignment = Alignment(
                horizontal="right", vertical="center"
            )
            worksheet.cell(row_number, 6).number_format = "#,##0.00"

            estimated_lines = max(1, math.ceil(len(record.service_name) / 54))
            worksheet.row_dimensions[row_number].height = min(18 * estimated_lines, 54)

    def _format_layout(self, worksheet: object, last_data_row: int) -> None:
        widths = {"A": 30, "B": 15, "C": 58, "D": 16, "E": 14, "F": 19}
        for column, width in widths.items():
            worksheet.column_dimensions[column].width = width

        worksheet.auto_filter.ref = f"A5:F{last_data_row}"
        worksheet.print_area = f"A1:F{last_data_row}"
        worksheet.print_title_rows = "1:5"
        worksheet.page_setup.orientation = "landscape"
        worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25
        worksheet.page_margins.top = 0.45
        worksheet.page_margins.bottom = 0.45
        worksheet.oddFooter.center.text = "Страница &P из &N"
        worksheet.oddFooter.center.size = 9
        worksheet.oddFooter.center.color = self.MUTED


class InteractiveInput:
    @staticmethod
    def choose_source_file() -> Path:
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            filename = filedialog.askopenfilename(
                title="Выберите таблицу с услугами",
                filetypes=(
                    ("Excel-файлы", "*.xls *.xlsx *.xlsm"),
                    ("Все файлы", "*.*"),
                ),
            )
            root.destroy()
            if not filename:
                raise ReportError("Выбор файла отменён.")
            return Path(filename)
        except ReportError:
            raise
        except Exception:
            raw_path = input("Введите полный путь к Excel-файлу: ").strip().strip('"')
            if not raw_path:
                raise ReportError("Путь к файлу не указан.")
            return Path(raw_path)

    @staticmethod
    def ask_period() -> tuple[date, date]:
        while True:
            date_from = InteractiveInput._ask_date("Дата начала (ДД.ММ.ГГГГ): ")
            date_to = InteractiveInput._ask_date("Дата окончания (ДД.ММ.ГГГГ): ")
            if date_from <= date_to:
                return date_from, date_to
            print("Дата начала не может быть позже даты окончания. Повторите ввод.\n")

    @staticmethod
    def _ask_date(prompt: str) -> date:
        while True:
            raw_value = input(prompt).strip()
            try:
                return parse_date(raw_value)
            except ValueError:
                print("Неверный формат. Пример корректной даты: 30.08.2026")


class ReportApplication:
    def __init__(self, config: AppConfig | None = None) -> None:
        self.config = config or AppConfig()
        self.reader = SourceReader(self.config)
        self.writer = ReportWorkbookWriter()

    def run(self, args: argparse.Namespace) -> list[Path]:
        source_path = Path(args.input).expanduser() if args.input else InteractiveInput.choose_source_file()
        source_path = source_path.resolve()
        self._validate_source(source_path)

        if args.date_from or args.date_to:
            if not (args.date_from and args.date_to):
                raise ReportError("Укажите одновременно --date-from и --date-to.")
            date_from = parse_date(args.date_from)
            date_to = parse_date(args.date_to)
            if date_from > date_to:
                raise ReportError("Дата начала не может быть позже даты окончания.")
        else:
            date_from, date_to = InteractiveInput.ask_period()

        print("\nЧитаю исходную таблицу и объединяю одинаковые услуги...")
        departments, stats, source_sheet = self.reader.aggregate(
            source_path,
            date_from,
            date_to,
            sheet_name=args.sheet,
            include_all_departments=args.all_departments,
        )
        if not departments:
            scope = "подразделениям" if args.all_departments else "женским консультациям"
            raise ReportError(
                f"За выбранный период не найдено строк по {scope}. "
                "Проверьте даты и значения в колонке «Подразделение»."
            )

        output_parent = (
            Path(args.output_dir).expanduser().resolve()
            if args.output_dir
            else source_path.parent
        )
        output_dir = OutputDirectoryFactory.create(output_parent, date_from, date_to)
        print(f"Найдено подразделений: {len(departments)}. Создаю отчёты...")
        files = self.writer.write_all(
            departments,
            output_dir,
            source_path,
            source_sheet,
            date_from,
            date_to,
        )
        self._print_result(files, stats, output_dir)
        return files

    @staticmethod
    def _validate_source(source_path: Path) -> None:
        if not source_path.exists() or not source_path.is_file():
            raise ReportError(f"Файл не найден: {source_path}")
        if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ReportError("Поддерживаются исходные файлы .xls, .xlsx и .xlsm.")

    @staticmethod
    def _print_result(
        files: Sequence[Path], stats: ProcessingStats, output_dir: Path
    ) -> None:
        print("\nГотово.")
        print(f"Создано файлов: {len(files)}")
        print(f"Папка: {output_dir}")
        print(f"Строк исходной таблицы просмотрено: {stats.source_rows:,}".replace(",", " "))
        print(f"Строк в выбранном периоде: {stats.rows_in_period:,}".replace(",", " "))
        print(
            f"Строк по выбранным подразделениям: "
            f"{stats.rows_in_selected_departments:,}".replace(",", " ")
        )
        print(f"Строк после объединения: {stats.aggregated_rows:,}".replace(",", " "))
        if stats.blank_quantity_rows:
            print(
                "Внимание: для строк с пустым количеством использовано значение 1: "
                f"{stats.blank_quantity_rows}"
            )
        if stats.skipped_rows:
            print(f"Пропущено строк с ошибками/пустыми обязательными полями: {stats.skipped_rows}")
            for warning in stats.warning_examples:
                print(f"  - {warning}")


def sanitize_filename(value: str, max_length: int = 125) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", clean_text(value))
    text = re.sub(r"\s+", " ", text).strip(" .")
    text = text[:max_length].rstrip(" .")
    return text or "Подразделение"


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Создание отдельных Excel-отчётов по услугам женских консультаций."
    )
    parser.add_argument("--input", help="Путь к исходному .xlsx/.xlsm")
    parser.add_argument("--date-from", help="Начальная дата, например 30.08.2026")
    parser.add_argument("--date-to", help="Конечная дата, например 01.09.2026")
    parser.add_argument("--output-dir", help="Папка, в которой создать каталог с отчётами")
    parser.add_argument("--sheet", help="Имя листа; по умолчанию ищется лист «услуги»")
    parser.add_argument(
        "--all-departments",
        action="store_true",
        help="Создать отчёты по всем подразделениям, а не только по ЖК",
    )
    parser.add_argument(
        "--no-pause",
        action="store_true",
        help="Не ждать Enter после завершения (удобно для автоматического запуска)",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    exit_code = 0
    try:
        ReportApplication().run(args)
    except (ReportError, PermissionError) as exc:
        exit_code = 1
        print(f"\nОшибка: {exc}", file=sys.stderr)
    except KeyboardInterrupt:
        exit_code = 130
        print("\nОперация отменена пользователем.", file=sys.stderr)
    except Exception as exc:
        exit_code = 1
        print(f"\nНепредвиденная ошибка: {exc}", file=sys.stderr)

    if not args.no_pause and sys.stdin.isatty():
        try:
            input("\nНажмите Enter, чтобы закрыть программу...")
        except (EOFError, KeyboardInterrupt):
            pass
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())