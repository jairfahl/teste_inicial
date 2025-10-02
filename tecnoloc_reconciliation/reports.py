"""Report generation and export utilities."""

from __future__ import annotations

from pathlib import Path
from statistics import mean
from typing import Dict, Iterable, List, Mapping

import pandas as pd

from .models import ErpRecord, PayfyExpense, ReconciliationResult


def _serialize_payfy(expense: PayfyExpense, *, origin: str | None = None) -> Dict[str, str]:
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
    if origin:
        payload = {"Origem": origin, **payload}
    return payload


def _serialize_erp(record: ErpRecord, *, origin: str | None = None) -> Dict[str, str]:
    payload = {
        "Usuário": record.user,
        "Data": record.date.strftime("%d/%m/%Y %H:%M"),
        "Valor": f"{record.value:.2f}",
        "Tipo": record.erp_type,
        "Match": record.match_type or "",
        "Motivo": record.failure_reason or "",
    }
    if origin:
        payload = {"Origem": origin, **payload}
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


def build_export_payload(result: ReconciliationResult) -> Dict[str, List[Dict[str, str]]]:
    """Prepare the export structure expected by :func:`export_excel`."""

    conciliados: List[Dict[str, str]] = []
    conciliados.extend(_serialize_payfy(item, origin="PayFy") for item in result.matched_payfy)
    conciliados.extend(_serialize_erp(item, origin="ERP") for item in result.matched_erp)

    nao_conciliados_cartao = [_serialize_erp(item) for item in result.unmatched_erp]
    nao_conciliados_despesas = [_serialize_payfy(item) for item in result.unmatched_payfy]

    summary = [{"Indicador": key, "Valor": value} for key, value in render_summary(result).items()]

    divergencias = [
        _serialize_payfy(item, origin="PayFy")
        for item in result.matched_payfy + result.unmatched_payfy
        if item.failure_reason and "diverg" in item.failure_reason.lower()
    ]
    divergencias.extend(
        _serialize_erp(item, origin="ERP")
        for item in result.matched_erp + result.unmatched_erp
        if item.failure_reason and "diverg" in item.failure_reason.lower()
    )

    aprovacao_fora_mes = [
        _serialize_payfy(item, origin="PayFy")
        for item in result.matched_payfy + result.unmatched_payfy
        if item.failure_reason == "Aprovação fora do mês corrente"
    ]

    return {
        "conciliados": conciliados,
        "nao conciliados - cartao": nao_conciliados_cartao,
        "nao conciliados - despesas": nao_conciliados_despesas,
        "resumo executivo": summary,
        "divergencia - valor": divergencias,
        "aprovacao fora do mes": aprovacao_fora_mes,
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


def export_excel(result_dict: Mapping[str, List[Dict[str, str]]], out_path: str | Path) -> Path:
    """Persist the reconciliation outcome into an Excel workbook."""

    output_path = Path(out_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, rows in result_dict.items():
            frame = pd.DataFrame(list(rows)) if rows else pd.DataFrame()
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output_path
