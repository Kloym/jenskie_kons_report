"""Факт ДОРЗ из услуг: один поток чтения, только обезличенные суммы.

Месяц заменяется целиком, а не прибавляется к ручному факту.
Сохранённый annual_plan.json никогда не изменяется этим модулем.
"""
from collections import Counter
from datetime import date

from zhk_annual_plan import AnnualPlanError, snapshot, validate_config
from zhk_monthly_report import (
    clean_text, decimal_number, identify_department, normalize_key, service_code,
)

TARGET_CODES = ("001983", "001984")


class MonthlyServiceActuals:
    """Собирает весь месяц независимо от фильтра начала основного отчёта."""

    def __init__(self, until: date):
        self.until = until
        self.start = until.replace(day=1)
        self.present = set()
        self.quantities = Counter()

    def observe(self, day, raw_department, raw_code, raw_quantity):
        if not self.start <= day <= self.until:
            return
        department = identify_department(raw_department, day) or clean_text(raw_department)
        if department:
            self.present.add(department)
        code = service_code(raw_code)
        if code not in TARGET_CODES:
            return
        quantity = decimal_number(raw_quantity, "Количество услуг ДОРЗ")
        if quantity < 0 or quantity != quantity.to_integral_value() or quantity > 10**9:
            raise AnnualPlanError("Количество услуг ДОРЗ должно быть целым числом от 0 до 1 000 000 000.")
        self.quantities[department, code] += int(quantity)

    def export(self):
        return {
            "from": self.start.isoformat(), "requestedAsOf": self.until.isoformat(),
            "presentDepartments": sorted(self.present),
            "rows": [{"department": department, "code": code, "quantity": value}
                     for (department, code), value in sorted(self.quantities.items())],
        }


def snapshot_from_services(config, data):
    """Создаёт проверенный годовой снимок без записи в файл настроек."""
    config = validate_config(config)
    actuals = data.get("serviceActuals")
    if not actuals:
        raise AnnualPlanError("Сначала прочитайте Excel в режиме «Факт месяца из Excel».")
    requested = date.fromisoformat(actuals["requestedAsOf"])
    first = requested.replace(day=1)
    configured = date.fromisoformat(config["asOf"])
    if (configured.year, configured.month) != (requested.year, requested.month):
        raise AnnualPlanError(
            "Месяц годового плана и выбранной даты «По» не совпадает. "
            "Обновите дату плана, план месяца и факт до месяца в окне «Годовой план…».")
    meta = data["meta"]
    if date.fromisoformat(meta["sourceStart"]) > first:
        raise AnnualPlanError("Для автоматического факта нужна полная выгрузка с 1-го числа месяца.")
    as_of = min(requested, date.fromisoformat(meta["sourceEnd"]))
    if as_of < first:
        raise AnnualPlanError("В Excel нет данных за месяц годового плана.")

    def key(name):
        return normalize_key(identify_department(name, as_of) or name)

    planned = {}
    for row in config["departments"]:
        name_key = key(row["name"])
        if name_key in planned:
            raise AnnualPlanError("Два подразделения плана соответствуют одной консультации. Проверьте названия.")
        planned[name_key] = row
    present = {key(name) for name in actuals["presentDepartments"]}
    missing = [row["name"] for name_key, row in planned.items() if name_key not in present]
    if missing:
        raise AnnualPlanError("В выгрузке за месяц отсутствуют подразделения плана: "
                              + ", ".join(missing) + ". Их факт нельзя считать нулём.")
    quantities, breakdown = Counter(), Counter()
    unmatched = set()
    for item in actuals["rows"]:
        name_key = key(item["department"])
        quantity = item["quantity"]
        if name_key not in planned and quantity:
            unmatched.add(item["department"] or "Подразделение не указано")
        quantities[name_key] += quantity
        breakdown[item["code"]] += quantity
    if unmatched:
        raise AnnualPlanError("Услуги ДОРЗ найдены вне подразделений годового плана: "
                              + ", ".join(sorted(unmatched)) + ". Проверьте соответствие подразделений.")
    for name_key, row in planned.items():
        row["monthActual"] = quantities[name_key]
    config["asOf"] = as_of.isoformat()
    result = snapshot(config)
    if as_of < requested:
        result["warnings"].append(
            f"Выбрано по {requested:%d.%m.%Y}, но последняя дата в Excel — {as_of:%d.%m.%Y}. "
            "Факт и план на дату рассчитаны по последней доступной дате.")
    result["actualSource"] = {
        "kind": "excel-services", "file": meta["source"], "sheet": meta["sheet"],
        "sha256": meta["sha256"], "codes": list(TARGET_CODES),
        "from": first.isoformat(), "asOf": as_of.isoformat(),
        "requestedAsOf": requested.isoformat(), "quantityByCode": dict(breakdown),
        "previousActualSource": "manual", "unit": "services",
    }
    return result
