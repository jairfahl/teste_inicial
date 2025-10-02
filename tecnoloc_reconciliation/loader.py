"""Data loading utilities for spreadsheet inputs."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd

from .models import ErpRecord, PayfyExpense

DATE_FORMATS = [
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
]


class DataValidationError(RuntimeError):
    """Raised when the input file is invalid."""


def _parse_date(value: str, *, field_name: str) -> datetime:
    text = value.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise DataValidationError(f"Data inválida '{value}' no campo '{field_name}'.")


def _parse_float(value: str, *, field_name: str) -> float:
    normalized = value.replace("R$", "").replace("$", "").replace(" ", "")
    normalized = (
        normalized.replace(".", "").replace(",", ".")
        if "," in normalized and normalized.count(",") == 1
        else normalized
    )
    try:
        return float(normalized)
    except ValueError as exc:
        raise DataValidationError(f"Valor inválido '{value}' no campo '{field_name}'.") from exc


def _read_rows(path: Path) -> Iterable[Dict[str, str]]:
    suffix = path.suffix.lower()

    if suffix == ".xlsx":
        try:
            frame = pd.read_excel(path, dtype=str, engine="openpyxl")
        except Exception as exc:  # pragma: no cover - dependente do ambiente
            raise DataValidationError(f"Não foi possível abrir a planilha: {path}") from exc
    elif suffix == ".csv":
        try:
            frame = pd.read_csv(path, dtype=str)
        except Exception as exc:  # pragma: no cover - dependente do ambiente
            raise DataValidationError(f"Não foi possível abrir o arquivo CSV: {path}") from exc
    else:
        raise DataValidationError(
            "Formato de arquivo não suportado. Utilize arquivos .xlsx ou .csv."
        )

    if frame.empty:
        raise DataValidationError("Planilha vazia.")

    frame.columns = ["" if col is None else str(col).strip() for col in frame.columns]

    if any(header == "" for header in frame.columns):
        raise DataValidationError("Cabeçalho inválido ou ausente na planilha.")

    frame = frame.fillna("")

    data_rows: List[Dict[str, str]] = []
    for _, series in frame.iterrows():
        row_dict: Dict[str, str] = {}
        for header, value in series.items():
            if header == "":
                continue
            if isinstance(value, str):
                row_dict[header] = value.strip()
            else:
                row_dict[header] = str(value)
        data_rows.append(row_dict)

    return data_rows


def load_payfy_expenses(path: Path) -> List[PayfyExpense]:
    """Load the detailed PayFy expenses file."""

    required = {"Usuário", "Data Transação", "Valor", "Status da Nota", "Categoria", "ID"}
    expenses: List[PayfyExpense] = []
    for row in _read_rows(path):
        if not required.issubset(row):
            missing = ", ".join(sorted(required - set(row)))
            raise DataValidationError(f"Campos obrigatórios ausentes: {missing}.")
        transaction_date = _parse_date(row["Data Transação"], field_name="Data Transação")
        approval = row.get("Data da Aprovação")
        approval_dt = _parse_date(approval, field_name="Data da Aprovação") if approval else None
        value = _parse_float(row["Valor"], field_name="Valor")
        expense = PayfyExpense(
            user=row["Usuário"],
            date=transaction_date,
            value=value,
            status=row.get("Status da Nota", ""),
            category=row.get("Categoria", ""),
            expense_id=row.get("ID"),
            approval_date=approval_dt,
            raw=row,
        )
        expenses.append(expense)
    return expenses


def load_erp_records(path: Path) -> List[ErpRecord]:
    """Load the ERP balance report."""

    required = {
        "Data",
        "Usuário",
        "Carga Empresa",
        "Carga Cartão",
        "Descarga Cartão",
        "Tarifas",
        "Reembolsos",
        "Saldo Empresa",
    }
    records: List[ErpRecord] = []
    for row in _read_rows(path):
        if not required.issubset(row):
            missing = ", ".join(sorted(required - set(row)))
            raise DataValidationError(f"Campos obrigatórios ausentes: {missing}.")
        date = _parse_date(row["Data"], field_name="Data")
        value_fields = ["Carga Empresa", "Carga Cartão", "Descarga Cartão", "Tarifas", "Reembolsos"]
        for field in value_fields:
            amount = _parse_float(row[field], field_name=field)
            if amount == 0:
                continue
            entry_type = "Tarifa" if field == "Tarifas" else field
            records.append(
                ErpRecord(
                    user=row.get("Usuário", ""),
                    date=date,
                    value=amount,
                    erp_type=entry_type,
                    raw=row,
                )
            )
    return records


def load_payfy_card_summary(path: Path) -> List[Tuple[str, float, float]]:
    """Load the PayFy technicians card summary (team, initial/final balances)."""

    required = {"Time", "Saldo Inicial", "Saldo Final"}
    summary: List[Tuple[str, float, float]] = []
    for row in _read_rows(path):
        if not required.issubset(row):
            missing = ", ".join(sorted(required - set(row)))
            raise DataValidationError(f"Campos obrigatórios ausentes: {missing}.")
        team = row["Time"]
        initial_balance = _parse_float(row["Saldo Inicial"], field_name="Saldo Inicial")
        final_balance = _parse_float(row["Saldo Final"], field_name="Saldo Final")
        summary.append((team, initial_balance, final_balance))
    return summary
