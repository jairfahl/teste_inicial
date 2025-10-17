"""Utilities for loading PayFy and ERP spreadsheets."""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterator, List, Tuple

import pandas as pd

from .models import EntryType, ErpRecord, PayfyExpense


class DataValidationError(RuntimeError):
    """Raised when the input file is invalid."""


def _normalize_header(header: object) -> str:
    value = "" if header is None else str(header)
    value = value.replace("\ufeff", "")  # Remove UTF-8 BOM markers.
    value = " ".join(value.strip().split())
    return value


def _stringify(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return " ".join(value.replace("\xa0", " ").strip().split())
    if hasattr(value, "to_pydatetime"):
        return str(value.to_pydatetime())
    return str(value)


def _parse_date(value: object, *, field_name: str) -> datetime:
    if isinstance(value, datetime):
        return value
    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime()
    text = _stringify(value)
    if not text:
        raise DataValidationError(f"Data inválida '{value}' no campo '{field_name}'.")
    try:
        parsed = pd.to_datetime(text, dayfirst=True, errors="raise")
    except (TypeError, ValueError) as exc:
        raise DataValidationError(f"Data inválida '{value}' no campo '{field_name}'.") from exc
    if isinstance(parsed, pd.Timestamp):
        return parsed.to_pydatetime()
    if isinstance(parsed, datetime):
        return parsed
    raise DataValidationError(f"Data inválida '{value}' no campo '{field_name}'.")


def _parse_float(value: object, *, field_name: str) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    text = _stringify(value)
    if not text:
        return 0.0
    negative = False
    cleaned = text.replace("R$", "").replace("$", "")
    cleaned = cleaned.replace("\xa0", " ").strip()
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = cleaned[1:-1]
        negative = True
    cleaned = cleaned.replace("–", "-").replace("—", "-")
    if cleaned.endswith("-"):
        cleaned = cleaned[:-1]
        negative = True
    cleaned = cleaned.replace(" ", "")
    if "," in cleaned and cleaned.count(",") == 1:
        cleaned = cleaned.replace(".", "")
        cleaned = cleaned.replace(",", ".")
    if cleaned in {"", "-", "+", ".", ","}:
        return 0.0
    try:
        number = float(cleaned)
    except ValueError as exc:
        raise DataValidationError(f"Valor inválido '{value}' no campo '{field_name}'.") from exc
    if negative:
        number = -abs(number)
    return number


def _normalize_numeric_sign(field: str, value: object) -> str | object:
    lowered = field.casefold()
    if "débito" not in lowered and "debito" not in lowered and "crédito" not in lowered and "credito" not in lowered:
        return value
    text = _stringify(value)
    if not text:
        return ""
    amount = _parse_float(text, field_name=field)
    if "débito" in lowered or "debito" in lowered:
        amount = -abs(amount)
    else:
        amount = abs(amount)
    return str(amount)


def _clean_row(row: Dict[str, object]) -> Dict[str, object]:
    cleaned: Dict[str, object] = {}
    empty = True
    for header, value in row.items():
        normalized_value = _normalize_numeric_sign(header, value)
        if normalized_value is value:
            if value is None or (isinstance(value, float) and pd.isna(value)):
                normalized_value = ""
            elif isinstance(value, datetime):
                normalized_value = value
            elif hasattr(value, "to_pydatetime"):
                normalized_value = value.to_pydatetime()
            else:
                normalized_value = _stringify(value)
        if normalized_value not in ("", None):
            empty = False
        cleaned[header] = normalized_value
    return cleaned if not empty else {}


def _prepare_raw(row: Dict[str, object]) -> Dict[str, object]:
    return {
        key: value if isinstance(value, datetime) else _stringify(value)
        for key, value in row.items()
    }


def _read_rows(path: Path) -> Iterator[Dict[str, object]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with path.open("r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            if not reader.fieldnames:
                raise DataValidationError("Arquivo CSV sem cabeçalho válido.")
            header_map = {name: _normalize_header(name) for name in reader.fieldnames}
            for raw_row in reader:
                normalized_row = {
                    header_map[key]: raw_row.get(key, "")
                    for key in reader.fieldnames
                }
                cleaned = _clean_row(normalized_row)
                if cleaned:
                    yield cleaned
        return
    if suffix in {".xlsx", ".xlsm"}:
        try:
            dataframe = pd.read_excel(path, dtype=object, engine="openpyxl")
        except ImportError as exc:
            raise DataValidationError(
                "Leitura de planilhas .xlsx requer a dependência 'openpyxl'."
            ) from exc
        except ValueError as exc:
            raise DataValidationError(f"Não foi possível ler a planilha: {exc}") from exc
        columns = [_normalize_header(name) for name in dataframe.columns]
        for _, series in dataframe.iterrows():
            normalized_row = {column: series[name] for column, name in zip(columns, dataframe.columns)}
            cleaned = _clean_row(normalized_row)
            if cleaned:
                yield cleaned
        return
    if suffix == ".xls":
        raise DataValidationError("Planilhas .xls não são suportadas; utilize o formato .xlsx.")
    raise DataValidationError(f"Formato de arquivo não suportado: {path.suffix}.")


def load_payfy_expenses(path: Path) -> List[PayfyExpense]:
    """Load the detailed PayFy expenses file."""

    required = {"Usuário", "Data Transação", "Valor", "Status da Nota", "Categoria", "ID"}
    expenses: List[PayfyExpense] = []
    for row in _read_rows(path):
        if not required.issubset(row):
            missing = ", ".join(sorted(required - set(row)))
            raise DataValidationError(f"Campos obrigatórios ausentes: {missing}.")
        transaction_date = _parse_date(row["Data Transação"], field_name="Data Transação")
        approval_raw = row.get("Data da Aprovação")
        approval_dt = _parse_date(approval_raw, field_name="Data da Aprovação") if approval_raw else transaction_date
        value = _parse_float(row["Valor"], field_name="Valor")
        expense = PayfyExpense(
            user=_stringify(row["Usuário"]),
            date=transaction_date,
            value=value,
            status=_stringify(row.get("Status da Nota", "")),
            category=_stringify(row.get("Categoria", "")),
            expense_id=_stringify(row.get("ID")) or None,
            approval_date=approval_dt,
            raw=_prepare_raw(row),
            entry_type=EntryType.DEBIT if value < 0 else EntryType.CREDIT,
        )
        expenses.append(expense)
    return expenses


def load_payfy_card_summary(path: Path) -> List[Tuple[str, float, float, float]]:
    """Load the PayFy technicians card summary."""

    required = {"Time", "Carga/Descarga do Cartão", "Tarifas", "Saldo"}
    summary: List[Tuple[str, float, float, float]] = []
    for row in _read_rows(path):
        if not required.issubset(row):
            missing = ", ".join(sorted(required - set(row)))
            raise DataValidationError(f"Campos obrigatórios ausentes: {missing}.")
        team = _stringify(row["Time"])
        card_flow = _parse_float(row["Carga/Descarga do Cartão"], field_name="Carga/Descarga do Cartão")
        fees = _parse_float(row["Tarifas"], field_name="Tarifas")
        balance = _parse_float(row["Saldo"], field_name="Saldo")
        summary.append((team, card_flow, fees, balance))
    return summary


def load_erp_records(path: Path) -> List[ErpRecord]:
    """Load the ERP balance report."""

    required = {"Data", "Usuário", "Carga/Descarga do Cartão", "Tarifas", "Saldo"}
    records: List[ErpRecord] = []
    for row in _read_rows(path):
        if not required.issubset(row):
            missing = ", ".join(sorted(required - set(row)))
            raise DataValidationError(f"Campos obrigatórios ausentes: {missing}.")
        date = _parse_date(row["Data"], field_name="Data")
        user = _stringify(row["Usuário"])
        _parse_float(row["Saldo"], field_name="Saldo")
        value_fields = {
            "Carga/Descarga do Cartão": "Carga/Descarga do Cartão",
            "Tarifas": "Tarifas",
            "Reembolsos": "Reembolsos",
        }
        for field, label in value_fields.items():
            if field not in row:
                continue
            amount = _parse_float(row[field], field_name=field)
            if amount == 0:
                continue
            record = ErpRecord(
                user=user,
                date=date,
                value=amount,
                erp_type=label,
                raw=_prepare_raw(row),
                entry_type=EntryType.DEBIT if amount < 0 else EntryType.CREDIT,
            )
            records.append(record)
    return records


def load_protheus_movements(path: Path) -> List[ErpRecord]:
    """Load the detailed ERP (Protheus) movement file."""

    required = {"data_mov", "valor_mov", "id_doc"}
    records: List[ErpRecord] = []
    for row in _read_rows(path):
        if not required.issubset(row):
            missing = ", ".join(sorted(required - set(row)))
            raise DataValidationError(f"Campos obrigatórios ausentes: {missing}.")
        date = _parse_date(row["data_mov"], field_name="data_mov")
        amount = _parse_float(row["valor_mov"], field_name="valor_mov")
        if amount == 0:
            continue
        user = (
            _stringify(row.get("usuario"))
            or _stringify(row.get("Usuário"))
            or _stringify(row.get("user"))
        )
        if not user:
            raise DataValidationError("Campo 'usuario' ausente no arquivo Protheus.")
        record = ErpRecord(
            user=user,
            date=date,
            value=amount,
            erp_type="Protheus",
            reference=_stringify(row.get("id_doc")) or None,
            raw=_prepare_raw(row),
            entry_type=EntryType.DEBIT if amount < 0 else EntryType.CREDIT,
        )
        records.append(record)
    return records
