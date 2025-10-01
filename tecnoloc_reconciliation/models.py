"""Data models used by the Tecnoloc reconciliation pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Dict, List, Optional


class EntryType(str, Enum):
    """Enumerates the debit/credit directions."""

    DEBIT = "DEBIT"
    CREDIT = "CREDIT"


@dataclass
class BaseRecord:
    """Base record for PayFy and ERP entries."""

    user: str
    date: datetime
    value: float
    raw: Dict[str, str] = field(default_factory=dict)
    entry_type: EntryType = EntryType.CREDIT


@dataclass
class PayfyExpense(BaseRecord):
    """Represents an expense as exported from PayFy."""

    status: str = ""
    category: str = ""
    expense_id: Optional[str] = None
    approval_date: Optional[datetime] = None
    match_id: Optional[str] = None
    match_type: Optional[str] = None
    failure_reason: Optional[str] = None


@dataclass
class ErpRecord(BaseRecord):
    """Represents an ERP Protheus record."""

    erp_type: str = ""
    reference: Optional[str] = None
    match_id: Optional[str] = None
    match_type: Optional[str] = None
    failure_reason: Optional[str] = None


@dataclass
class ReconciliationResult:
    """Aggregated reconciliation output."""

    matched_payfy: List[PayfyExpense] = field(default_factory=list)
    matched_erp: List[ErpRecord] = field(default_factory=list)
    unmatched_payfy: List[PayfyExpense] = field(default_factory=list)
    unmatched_erp: List[ErpRecord] = field(default_factory=list)
    diagnostics: Dict[str, int] = field(default_factory=dict)
