"""Reconciliation matching strategies."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional

from .models import ErpRecord, PayfyExpense


def _set_failure(record: PayfyExpense | ErpRecord, reason: str) -> None:
    record.failure_reason = reason
    if hasattr(record, "failure_cause"):
        record.failure_cause = reason


@dataclass
class Match:
    payfy_items: List[PayfyExpense]
    erp_items: List[ErpRecord]
    match_type: str
    tolerance: float = 0.0


def exact_match(payfy: List[PayfyExpense], erp: List[ErpRecord]) -> List[Match]:
    matches: List[Match] = []
    payfy_by_id: Dict[str, List[PayfyExpense]] = defaultdict(list)
    for expense in payfy:
        if expense.match_id or expense.failure_reason or not expense.expense_id:
            continue
        payfy_by_id[expense.expense_id].append(expense)

    for record in erp:
        if record.match_id or record.failure_reason or not record.document_id:
            continue
        candidates = payfy_by_id.get(record.document_id)
        if not candidates:
            continue
        movement_date = record.movement_date or record.date
        if not movement_date:
            continue
        matched_expense: Optional[PayfyExpense] = None
        for candidate in candidates:
            if candidate.match_id or candidate.failure_reason:
                continue
            if candidate.value != record.value:
                continue
            approval_date = candidate.approval_date
            if not approval_date:
                continue
            if (
                approval_date.month == movement_date.month
                and approval_date.year == movement_date.year
            ):
                matched_expense = candidate
                break
        if not matched_expense:
            continue
        candidates.remove(matched_expense)
        match = Match([matched_expense], [record], "Match exato (1:1)")
        matches.append(match)
        match_id = (
            record.document_id
            or record.reference
            or matched_expense.expense_id
            or "MATCH"
        )
        matched_expense.match_id = match_id
        matched_expense.match_type = match.match_type
        record.match_id = match_id
        record.match_type = match.match_type
    return matches


def classify_unmatched(payfy: Iterable[PayfyExpense], erp: Iterable[ErpRecord]) -> None:
    payfy_by_id: Dict[str, List[PayfyExpense]] = defaultdict(list)
    for expense in payfy:
        if expense.expense_id:
            payfy_by_id[expense.expense_id].append(expense)

    erp_by_id: Dict[str, List[ErpRecord]] = defaultdict(list)
    for record in erp:
        if record.document_id:
            erp_by_id[record.document_id].append(record)

    for expense in payfy:
        if expense.match_id or expense.failure_reason:
            continue
        if not expense.expense_id:
            _set_failure(expense, "Despesa sem identificador")
            continue
        candidates = erp_by_id.get(expense.expense_id)
        if not candidates:
            _set_failure(expense, "Sem correspondência no Protheus")
            continue
        same_value = [record for record in candidates if record.value == expense.value]
        if not same_value:
            _set_failure(expense, "Valor divergente no Protheus")
            continue
        approval_date = expense.approval_date
        if not approval_date:
            _set_failure(expense, "Despesa sem aprovação registrada")
            continue
        competence_match = False
        for record in same_value:
            movement_date = record.movement_date or record.date
            if not movement_date:
                continue
            if (
                approval_date.month == movement_date.month
                and approval_date.year == movement_date.year
            ):
                competence_match = True
                break
        if not competence_match:
            _set_failure(expense, "Aprovação fora do mês")

    for record in erp:
        if record.match_id or record.failure_reason:
            continue
        if not record.document_id:
            _set_failure(record, "Lançamento sem identificador")
            continue
        candidates = payfy_by_id.get(record.document_id)
        if not candidates:
            _set_failure(record, "Sem correspondência no PayFy")
            continue
        same_value = [expense for expense in candidates if expense.value == record.value]
        if not same_value:
            _set_failure(record, "Valor divergente no PayFy")
            continue
        movement_date = record.movement_date or record.date
        if not movement_date:
            _set_failure(record, "Competência ausente no Protheus")
            continue
        competence_match = False
        for expense in same_value:
            approval_date = expense.approval_date
            if not approval_date:
                continue
            if (
                approval_date.month == movement_date.month
                and approval_date.year == movement_date.year
            ):
                competence_match = True
                break
        if not competence_match:
            _set_failure(record, "Competência divergente no PayFy")


def reconcile(payfy: List[PayfyExpense], erp: List[ErpRecord]) -> List[Match]:
    matches = []
    matches.extend(exact_match(payfy, erp))
    classify_unmatched(payfy, erp)
    return matches
