"""Reconciliation matching strategies."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

from .models import ErpRecord, PayfyExpense


@dataclass
class Match:
    payfy_items: List[PayfyExpense]
    erp_items: List[ErpRecord]
    match_type: str
    tolerance: float = 0.0


def _within_tolerance(a: float, b: float, tolerance: float) -> bool:
    return abs(a - b) <= tolerance


def _date_close(a: datetime, b: datetime, tolerance_days: int = 1) -> bool:
    return abs((a - b).days) <= tolerance_days


def exact_match(payfy: List[PayfyExpense], erp: List[ErpRecord]) -> List[Match]:
    matches: List[Match] = []
    payfy_by_key: Dict[Tuple[str, datetime, float], List[PayfyExpense]] = defaultdict(list)
    for item in payfy:
        if item.match_id or item.failure_reason:
            continue
        payfy_by_key[(item.user, item.date, item.value)].append(item)

    for record in erp:
        if record.match_id or record.failure_reason:
            continue
        key = (record.user, record.date, record.value)
        candidates = payfy_by_key.get(key)
        if candidates:
            expense = candidates.pop(0)
            match = Match([expense], [record], "Match exato (1:1)")
            matches.append(match)
            expense.match_id = record.reference or expense.expense_id or "MATCH"
            record.match_id = expense.expense_id or record.reference or "MATCH"
            expense.match_type = match.match_type
            record.match_type = match.match_type
    return matches


def tolerance_match(payfy: List[PayfyExpense], erp: List[ErpRecord], tolerance: float = 0.01) -> List[Match]:
    matches: List[Match] = []
    for record in erp:
        if record.match_id or record.failure_reason:
            continue
        for expense in payfy:
            if expense.match_id or expense.failure_reason:
                continue
            if expense.user != record.user:
                continue
            if not _date_close(expense.date, record.date):
                continue
            if _within_tolerance(expense.value, record.value, tolerance):
                match = Match([expense], [record], "Match tolerância (1:1)", tolerance=tolerance)
                matches.append(match)
                expense.match_id = record.reference or expense.expense_id or "MATCH"
                record.match_id = expense.expense_id or record.reference or "MATCH"
                expense.match_type = match.match_type
                record.match_type = match.match_type
                break
    return matches


def aggregation_match(payfy: List[PayfyExpense], erp: List[ErpRecord], tolerance: float = 0.01) -> List[Match]:
    matches: List[Match] = []
    unresolved_payfy = [p for p in payfy if not p.match_id and not p.failure_reason]
    unresolved_erp = [e for e in erp if not e.match_id and not e.failure_reason]

    for record in unresolved_erp:
        bucket = [expense for expense in unresolved_payfy if expense.user == record.user and _date_close(expense.date, record.date)]
        bucket.sort(key=lambda item: item.value, reverse=True)
        combination: List[PayfyExpense] = []
        total = 0.0
        for expense in bucket:
            if expense in combination:
                continue
            if total + expense.value <= record.value + tolerance:
                combination.append(expense)
                total += expense.value
            if _within_tolerance(total, record.value, tolerance):
                match = Match(combination.copy(), [record], "Match agregado (N:1)", tolerance=tolerance)
                matches.append(match)
                match_id = record.reference or "AGG"
                for item in combination:
                    item.match_id = match_id
                    item.match_type = match.match_type
                record.match_id = match_id
                record.match_type = match.match_type
                break
    return matches


def classify_unmatched(payfy: Iterable[PayfyExpense], erp: Iterable[ErpRecord]) -> None:
    for expense in payfy:
        if expense.match_id:
            continue
        if not expense.failure_reason:
            expense.failure_reason = "Sem correspondência entre bases"
    for record in erp:
        if record.match_id:
            continue
        if not record.failure_reason:
            record.failure_reason = "Sem correspondência entre bases"


def reconcile(payfy: List[PayfyExpense], erp: List[ErpRecord]) -> List[Match]:
    matches = []
    matches.extend(exact_match(payfy, erp))
    matches.extend(tolerance_match(payfy, erp))
    matches.extend(aggregation_match(payfy, erp))
    classify_unmatched(payfy, erp)
    return matches
