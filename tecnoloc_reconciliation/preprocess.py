"""Pre-processing steps for reconciliation inputs."""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime, timedelta
from typing import Dict, Iterable, List, Tuple

from .models import EntryType, ErpRecord, PayfyExpense

PAYFY_CATEGORY_MAP = {
    "Hospedagem": "Desp. Viagem – Hospedagem",
    "Alimentação": "Desp. Viagem – Alimentação",
    "Combustível": "Desp. Operacional – Frota",
    "Pedágio": "Desp. Operacional – Frota",
}


def _set_failure(record: PayfyExpense | ErpRecord, reason: str) -> None:
    record.failure_reason = reason
    if hasattr(record, "failure_cause"):
        record.failure_cause = reason


class PeriodNotProvidedError(RuntimeError):
    """Raised when the reconciliation period is missing."""


def normalize_entry_types(records: Iterable[PayfyExpense | ErpRecord]) -> None:
    for record in records:
        record.entry_type = EntryType.DEBIT if record.value < 0 else EntryType.CREDIT
        record.value = abs(record.value)


def map_categories(expenses: Iterable[PayfyExpense]) -> None:
    for expense in expenses:
        normalized = expense.category.strip()
        expense.category = PAYFY_CATEGORY_MAP.get(normalized, normalized or "Revisão manual")


def apply_date_rules(expenses: Iterable[PayfyExpense]) -> None:
    for expense in expenses:
        original_date = expense.date
        if original_date.hour > 18 or (original_date.hour == 18 and original_date.minute >= 1):
            expense.date = (original_date + timedelta(days=1)).replace(hour=original_date.hour, minute=original_date.minute)
        if expense.approval_date and abs((expense.approval_date - expense.date).days) <= 1:
            continue
        if expense.approval_date and expense.approval_date.month != expense.date.month:
            _set_failure(expense, "Aprovação fora do mês")


def detect_duplicates(expenses: List[PayfyExpense]) -> Dict[str, List[PayfyExpense]]:
    registry: Dict[Tuple[str, datetime, float], List[PayfyExpense]] = defaultdict(list)
    for expense in expenses:
        registry[(expense.user, expense.date, expense.value)].append(expense)
    duplicates = {key: entries for key, entries in registry.items() if len(entries) > 1}
    for entries in duplicates.values():
        for entry in entries:
            _set_failure(entry, "Duplicidade detectada")
    return duplicates


def ensure_status_validated(expenses: Iterable[PayfyExpense]) -> None:
    for expense in expenses:
        status = expense.status.lower()
        if status != "validado" and not expense.approval_date:
            _set_failure(expense, "Status não validado")
        elif status == "validado" and expense.approval_date is None:
            expense.approval_date = expense.date


def validate_period(expenses: Iterable[PayfyExpense], reference: datetime) -> None:
    for expense in expenses:
        if expense.date.month != reference.month or expense.date.year != reference.year:
            _set_failure(expense, "Transação fora do período")


def summarize_failures(expenses: Iterable[PayfyExpense], erp_records: Iterable[ErpRecord]) -> Dict[str, int]:
    counter: Counter[str] = Counter()
    for expense in expenses:
        reason = expense.failure_cause or expense.failure_reason
        if reason:
            counter[reason] += 1
    for record in erp_records:
        reason = record.failure_cause or record.failure_reason
        if reason:
            counter[reason] += 1
    return dict(counter)
