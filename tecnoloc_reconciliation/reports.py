"""Reporting utilities for the reconciliation pipeline."""

from __future__ import annotations

from pathlib import Path
from statistics import mean
from typing import Callable, Dict, Iterable, List, Tuple

from .models import ErpRecord, PayfyExpense, ReconciliationResult


def _serialize_payfy(expense: PayfyExpense) -> Dict[str, str]:
    payload = {
        "Usuário": expense.user,
        "Data": expense.date.strftime("%d/%m/%Y %H:%M"),
        "Valor": f"{expense.value:.2f}",
        "Status": expense.status,
        "Categoria": expense.category,
        "ID": expense.expense_id or "",
        "Match": expense.match_type or "",
        "Motivo": expense.failure_reason or "",
    }
    if expense.approval_date:
        payload["Aprovação"] = expense.approval_date.strftime("%d/%m/%Y %H:%M")
    else:
        payload["Aprovação"] = ""
    return payload


def _serialize_erp(record: ErpRecord) -> Dict[str, str]:
    payload = {
        "Usuário": record.user,
        "Data": record.date.strftime("%d/%m/%Y %H:%M"),
        "Valor": f"{record.value:.2f}",
        "Tipo": record.erp_type,
        "Match": record.match_type or "",
        "Motivo": record.failure_reason or "",
    }
    if record.reference:
        payload["Referência"] = record.reference
    return payload


def build_reconciliation_result(payfy: List[PayfyExpense], erp: List[ErpRecord], diagnostics: Dict[str, int]) -> ReconciliationResult:
    matched_payfy = [expense for expense in payfy if expense.match_id]
    unmatched_payfy = [expense for expense in payfy if not expense.match_id]
    matched_erp = [record for record in erp if record.match_id]
    unmatched_erp = [record for record in erp if not record.match_id]
    return ReconciliationResult(
        matched_payfy=matched_payfy,
        unmatched_payfy=unmatched_payfy,
        matched_erp=matched_erp,
        unmatched_erp=unmatched_erp,
        diagnostics=diagnostics,
    )


def render_summary(result: ReconciliationResult) -> Dict[str, str]:
    total_payfy = sum(item.value for item in result.matched_payfy + result.unmatched_payfy)
    total_erp = sum(item.value for item in result.matched_erp + result.unmatched_erp)
    automatic_rate = (
        len(result.matched_payfy) / len(result.matched_payfy + result.unmatched_payfy)
        if result.matched_payfy or result.unmatched_payfy
        else 0
    )
    uncategorized_rate = (
        len([item for item in result.matched_payfy + result.unmatched_payfy if item.category == "Revisão manual"])
        / len(result.matched_payfy + result.unmatched_payfy)
        if result.matched_payfy or result.unmatched_payfy
        else 0
    )
    approval_deltas = [
        (expense.approval_date - expense.date).days
        for expense in result.matched_payfy + result.unmatched_payfy
        if expense.approval_date
    ]
    avg_time = mean(approval_deltas) if approval_deltas else 0
    adjustments = sum(item.value for item in result.unmatched_erp if item.erp_type in {"Tarifa", "Reembolsos"})

    return {
        "Total PayFy": f"{total_payfy:.2f}",
        "Total ERP": f"{total_erp:.2f}",
        "% Conciliação Automática": f"{automatic_rate * 100:.1f}%",
        "% Despesas sem categoria": f"{uncategorized_rate * 100:.1f}%",
        "Tempo médio transação-aprovação": f"{avg_time:.1f} dias",
        "Ajustes manuais": f"{adjustments:.2f}",
    }


def render_table(title: str, records: Iterable[Dict[str, str]]) -> str:
    lines = [title]
    for record in records:
        serialized = ", ".join(f"{key}: {value}" for key, value in record.items())
        lines.append(f"- {serialized}")
    return "\n".join(lines)


def render_reports(result: ReconciliationResult) -> str:
    lines = ["# Relatório de Conciliação Tecnoloc"]
    lines.append("\n## Resumo Executivo")
    for key, value in render_summary(result).items():
        lines.append(f"- {key}: {value}")

    lines.append("\n## Diagnóstico Automático")
    for key, value in result.diagnostics.items():
        lines.append(f"- {key}: {value}")

    lines.append("\n## Despesas Conciliadas")
    lines.append(
        render_table(
            "### Relatório Conciliado",
            (_serialize_payfy(item) for item in result.matched_payfy),
        )
    )
    lines.append(
        render_table(
            "### Lançamentos ERP Conciliados",
            (_serialize_erp(item) for item in result.matched_erp),
        )
    )

    lines.append("\n## Pendências")
    lines.append(
        render_table(
            "### Despesas Não Conciliadas",
            (_serialize_payfy(item) for item in result.unmatched_payfy),
        )
    )
    lines.append(
        render_table(
            "### Registros ERP Não Conciliados",
            (_serialize_erp(item) for item in result.unmatched_erp),
        )
    )
    return "\n".join(lines)


def _ensure_openpyxl() -> Tuple["Workbook", Callable[[int], str]]:
    try:
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter
    except ModuleNotFoundError as exc:
        raise RuntimeError("A geração de relatórios Excel requer a instalação de 'openpyxl'.") from exc
    return Workbook, get_column_letter


def _autosize_columns(worksheet, get_column_letter) -> None:
    widths: Dict[int, int] = {}
    for row in worksheet.iter_rows(values_only=True):
        for index, value in enumerate(row, start=1):
            if value is None:
                continue
            text = str(value)
            widths[index] = max(widths.get(index, 0), len(text))
    for index, width in widths.items():
        worksheet.column_dimensions[get_column_letter(index)].width = min(width + 2, 80)


def _write_sheet(workbook, get_column_letter, title: str, rows: List[Dict[str, str]]) -> None:
    sheet = workbook.create_sheet(title)
    if not rows:
        return
    headers = list(rows[0].keys())
    sheet.append(headers)
    for row in rows:
        sheet.append([row.get(header, "") for header in headers])
    _autosize_columns(sheet, get_column_letter)


def write_excel_report(result: ReconciliationResult, path: Path) -> None:
    """Persist the reconciliation result into an Excel workbook."""

    Workbook, get_column_letter = _ensure_openpyxl()
    workbook = Workbook()
    workbook.remove(workbook.active)

    summary_sheet = workbook.create_sheet("Resumo")
    for key, value in render_summary(result).items():
        summary_sheet.append([key, value])
    _autosize_columns(summary_sheet, get_column_letter)

    diagnostics_sheet = workbook.create_sheet("Diagnóstico")
    if result.diagnostics:
        diagnostics_sheet.append(["Motivo", "Ocorrências"])
        for key, value in result.diagnostics.items():
            diagnostics_sheet.append([key, value])
    _autosize_columns(diagnostics_sheet, get_column_letter)

    payfy_matched = [_serialize_payfy(item) for item in result.matched_payfy]
    _write_sheet(workbook, get_column_letter, "PayFy Conciliado", payfy_matched)

    erp_matched = [_serialize_erp(item) for item in result.matched_erp]
    _write_sheet(workbook, get_column_letter, "ERP Conciliado", erp_matched)

    payfy_unmatched = [_serialize_payfy(item) for item in result.unmatched_payfy]
    _write_sheet(workbook, get_column_letter, "PayFy Pendente", payfy_unmatched)

    erp_unmatched = [_serialize_erp(item) for item in result.unmatched_erp]
    _write_sheet(workbook, get_column_letter, "ERP Pendente", erp_unmatched)

    workbook.save(path)

