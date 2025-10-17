"""Utilities to consolidate reconciliation results into spreadsheet reports."""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from statistics import mean
from typing import Dict, Iterable, List, Mapping, MutableMapping, Optional, Sequence

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from .models import ErpRecord, PayfyExpense, ReconciliationResult


SheetFrames = Mapping[str, pd.DataFrame]


@dataclass
class AggregatedFailure:
    """Represents aggregated information for a given failure reason."""

    reason: str
    count: int
    value: float
    origins: Sequence[str]
    cause_type: str
    action: str


# Mapping between failure reasons and categorized causes/actions.
FAILURE_CAUSE_MAP: Mapping[str, str] = {
    "Sem correspondência entre bases": "Integração de dados",
    "Duplicidade detectada": "Processo PayFy",
    "Status não validado": "Processo de aprovação",
    "Transação fora do período": "Gestão do período",
    "Aprovação fora do mês corrente": "Processo de aprovação",
}

FAILURE_ACTION_MAP: Mapping[str, str] = {
    "Sem correspondência entre bases": "Revisar importação e aplicar conciliações manuais.",
    "Duplicidade detectada": "Checar lançamentos repetidos e cancelar duplicatas.",
    "Status não validado": "Notificar responsáveis para validar notas pendentes.",
    "Transação fora do período": "Ajustar janela de lançamento e reforçar comunicação de datas.",
    "Aprovação fora do mês corrente": "Revisar SLAs de aprovação junto às lideranças.",
}


def _serialize_payfy(expense: PayfyExpense) -> Dict[str, str]:
    return {
        "Origem": "PayFy",
        "Usuário": expense.user,
        "Data": expense.date.strftime("%d/%m/%Y %H:%M"),
        "Valor": f"{expense.value:.2f}",
        "Status": expense.status or "",
        "Categoria": expense.category or "",
        "Tipo": "",
        "ID": expense.expense_id or "",
        "Match": expense.match_type or "",
        "Motivo": expense.failure_reason or "",
    }


def _serialize_erp(record: ErpRecord) -> Dict[str, str]:
    return {
        "Origem": "ERP",
        "Usuário": record.user,
        "Data": record.date.strftime("%d/%m/%Y %H:%M"),
        "Valor": f"{record.value:.2f}",
        "Status": "",
        "Categoria": "",
        "Tipo": record.erp_type or "",
        "ID": record.reference or "",
        "Match": record.match_type or "",
        "Motivo": record.failure_reason or "",
    }


def build_reconciliation_result(
    payfy: List[PayfyExpense],
    erp: List[ErpRecord],
    diagnostics: Dict[str, int],
) -> ReconciliationResult:
    """Create a :class:`ReconciliationResult` from processed records."""

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


def _format_currency(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _format_percent(value: float) -> str:
    return f"{value * 100:.1f}%"


def _build_conciliated_sheet(result: ReconciliationResult) -> pd.DataFrame:
    records: List[Dict[str, str]] = []
    records.extend(_serialize_payfy(item) for item in result.matched_payfy)
    records.extend(_serialize_erp(item) for item in result.matched_erp)
    return pd.DataFrame(records, columns=[
        "Origem",
        "Usuário",
        "Data",
        "Valor",
        "Status",
        "Categoria",
        "Tipo",
        "ID",
        "Match",
        "Motivo",
    ])


def _build_unreconciled_payfy_sheet(result: ReconciliationResult) -> pd.DataFrame:
    records = [_serialize_payfy(item) for item in result.unmatched_payfy]
    return pd.DataFrame(records, columns=[
        "Origem",
        "Usuário",
        "Data",
        "Valor",
        "Status",
        "Categoria",
        "Tipo",
        "ID",
        "Match",
        "Motivo",
    ])


def _build_unreconciled_erp_sheet(result: ReconciliationResult) -> pd.DataFrame:
    records = [_serialize_erp(item) for item in result.unmatched_erp]
    return pd.DataFrame(records, columns=[
        "Origem",
        "Usuário",
        "Data",
        "Valor",
        "Status",
        "Categoria",
        "Tipo",
        "ID",
        "Match",
        "Motivo",
    ])


def _build_summary_sheet(result: ReconciliationResult) -> pd.DataFrame:
    payfy_all = result.matched_payfy + result.unmatched_payfy
    erp_all = result.matched_erp + result.unmatched_erp

    total_payfy_value = sum(item.value for item in payfy_all)
    total_erp_value = sum(item.value for item in erp_all)
    total_payfy_count = len(payfy_all)
    total_erp_count = len(erp_all)
    matched_payfy_count = len(result.matched_payfy)
    matched_erp_count = len(result.matched_erp)
    unmatched_payfy_value = sum(item.value for item in result.unmatched_payfy)
    unmatched_erp_value = sum(item.value for item in result.unmatched_erp)

    approval_deltas = [
        (expense.approval_date - expense.date).days
        for expense in payfy_all
        if expense.approval_date
    ]
    average_approval_time = mean(approval_deltas) if approval_deltas else 0.0
    automatic_rate = matched_payfy_count / total_payfy_count if total_payfy_count else 0.0
    uncategorized_rate = (
        len([item for item in payfy_all if item.category == "Revisão manual"])
        / total_payfy_count
        if total_payfy_count
        else 0.0
    )

    indicators = [
        ("Volume de despesas PayFy", str(total_payfy_count)),
        ("Volume de lançamentos ERP", str(total_erp_count)),
        ("Despesas conciliadas", str(matched_payfy_count)),
        ("Lançamentos ERP conciliados", str(matched_erp_count)),
        ("Total despesas PayFy", _format_currency(total_payfy_value)),
        ("Total ERP", _format_currency(total_erp_value)),
        ("Valor pendente PayFy", _format_currency(unmatched_payfy_value)),
        ("Valor pendente ERP", _format_currency(unmatched_erp_value)),
        ("% Conciliação Automática", _format_percent(automatic_rate)),
        ("% Despesas sem categoria", _format_percent(uncategorized_rate)),
        (
            "Tempo médio transação-aprovação (dias)",
            f"{average_approval_time:.1f}",
        ),
    ]

    indicators.extend((f"Diagnóstico – {key}", str(value)) for key, value in result.diagnostics.items())

    return pd.DataFrame(indicators, columns=["Indicador", "Valor"])


def _build_duplicates_sheet(result: ReconciliationResult) -> pd.DataFrame:
    duplicates: MutableMapping[tuple[str, str, str], List[Dict[str, str]]] = defaultdict(list)
    for expense in result.matched_payfy + result.unmatched_payfy:
        if expense.failure_reason != "Duplicidade detectada":
            continue
        payload = _serialize_payfy(expense)
        key = (payload["Usuário"], payload["Data"], payload["Valor"])
        duplicates[key].append(payload)

    rows: List[Dict[str, str]] = []
    for index, entries in enumerate(duplicates.values(), start=1):
        group_id = f"DUP-{index:03d}"
        for payload in entries:
            augmented = dict(payload)
            augmented["Grupo"] = group_id
            rows.append(augmented)

    if not rows:
        return pd.DataFrame(columns=[
            "Grupo",
            "Origem",
            "Usuário",
            "Data",
            "Valor",
            "Status",
            "Categoria",
            "Tipo",
            "ID",
            "Match",
            "Motivo",
        ])

    columns = [
        "Grupo",
        "Origem",
        "Usuário",
        "Data",
        "Valor",
        "Status",
        "Categoria",
        "Tipo",
        "ID",
        "Match",
        "Motivo",
    ]
    return pd.DataFrame(rows, columns=columns)


def _aggregate_failures(result: ReconciliationResult) -> List[AggregatedFailure]:
    counter: Dict[str, Counter[str]] = defaultdict(Counter)
    value_accumulator: Dict[str, float] = defaultdict(float)

    for origin, records in (
        ("PayFy", result.matched_payfy + result.unmatched_payfy),
        ("ERP", result.matched_erp + result.unmatched_erp),
    ):
        for record in records:
            reason = record.failure_reason
            if not reason:
                continue
            counter[reason][origin] += 1
            value_accumulator[reason] += record.value

    aggregated: List[AggregatedFailure] = []
    for reason, origin_counts in counter.items():
        count = sum(origin_counts.values())
        origins = [f"{origin} ({origin_counts[origin]})" for origin in sorted(origin_counts)]
        cause_type = FAILURE_CAUSE_MAP.get(reason, "Outros")
        action = FAILURE_ACTION_MAP.get(reason, "Analisar causa raiz e definir tratativa com o time responsável.")
        aggregated.append(
            AggregatedFailure(
                reason=reason,
                count=count,
                value=value_accumulator[reason],
                origins=origins,
                cause_type=cause_type,
                action=action,
            )
        )

    aggregated.sort(key=lambda item: item.count, reverse=True)
    return aggregated


def _build_action_plan_sheet(result: ReconciliationResult) -> pd.DataFrame:
    aggregated = _aggregate_failures(result)
    rows = [
        {
            "Motivo": item.reason,
            "Tipo de Causa": item.cause_type,
            "Quantidade": item.count,
            "Valor Total": _format_currency(item.value),
            "Origem": ", ".join(item.origins),
            "Ação Sugerida": item.action,
        }
        for item in aggregated
    ]

    return pd.DataFrame(
        rows,
        columns=[
            "Motivo",
            "Tipo de Causa",
            "Quantidade",
            "Valor Total",
            "Origem",
            "Ação Sugerida",
        ],
    )


def _build_sheet_frames(result: ReconciliationResult) -> SheetFrames:
    frames: Dict[str, pd.DataFrame] = {
        "Conciliados": _build_conciliated_sheet(result),
        "Não Conciliados – Cartão": _build_unreconciled_erp_sheet(result),
        "Não Conciliados – Despesas": _build_unreconciled_payfy_sheet(result),
        "Resumo Executivo": _build_summary_sheet(result),
        "Duplicidades": _build_duplicates_sheet(result),
        "Plano de Ação": _build_action_plan_sheet(result),
    }
    return frames


def build_workbook(result: ReconciliationResult, frames: Optional[SheetFrames] = None) -> Workbook:
    """Create an :class:`openpyxl.Workbook` populated with the reconciliation data."""

    if frames is None:
        frames = _build_sheet_frames(result)
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    for sheet_name, frame in frames.items():
        worksheet = workbook.create_sheet(sheet_name[:31])
        if frame.empty:
            worksheet.append(list(frame.columns))
            continue
        for row in dataframe_to_rows(frame, index=False, header=True):
            worksheet.append(row)

    return workbook


def _sheet_metrics(frame: pd.DataFrame) -> Dict[str, int]:
    return {
        "rows": int(frame.shape[0]),
        "columns": int(frame.shape[1]),
    }


def build_metadata(result: ReconciliationResult, frames: SheetFrames) -> Dict[str, object]:
    """Compile useful metadata for logging/CLI consumption."""

    payfy_all = result.matched_payfy + result.unmatched_payfy
    erp_all = result.matched_erp + result.unmatched_erp

    metadata = {
        "matched_payfy": len(result.matched_payfy),
        "matched_erp": len(result.matched_erp),
        "unmatched_payfy": len(result.unmatched_payfy),
        "unmatched_erp": len(result.unmatched_erp),
        "total_payfy_value": sum(item.value for item in payfy_all),
        "total_erp_value": sum(item.value for item in erp_all),
        "sheets": {
            sheet: _sheet_metrics(frame)
            for sheet, frame in frames.items()
        },
    }
    return metadata


def write_report(result: ReconciliationResult, output_path: Path) -> Dict[str, object]:
    """Persist the reconciliation report to ``output_path`` and return metadata."""

    frames = _build_sheet_frames(result)
    workbook = build_workbook(result, frames)
    workbook.save(output_path)
    metadata = build_metadata(result, frames)
    metadata["path"] = str(output_path)
    return metadata


def render_summary(result: ReconciliationResult) -> Dict[str, str]:
    """Backward compatible textual summary used by legacy CLI reporting."""

    payfy_all = result.matched_payfy + result.unmatched_payfy
    erp_all = result.matched_erp + result.unmatched_erp

    total_payfy = sum(item.value for item in payfy_all)
    total_erp = sum(item.value for item in erp_all)
    automatic_rate = (
        len(result.matched_payfy) / len(payfy_all)
        if payfy_all
        else 0
    )
    uncategorized_rate = (
        len([item for item in payfy_all if item.category == "Revisão manual"]) / len(payfy_all)
        if payfy_all
        else 0
    )
    approval_deltas = [
        (expense.approval_date - expense.date).days
        for expense in payfy_all
        if expense.approval_date
    ]
    avg_time = mean(approval_deltas) if approval_deltas else 0
    adjustments = sum(
        item.value for item in result.unmatched_erp if item.erp_type in {"Tarifa", "Reembolsos"}
    )

    return {
        "Total PayFy": f"{total_payfy:.2f}",
        "Total ERP": f"{total_erp:.2f}",
        "% Conciliação Automática": f"{automatic_rate * 100:.1f}%",
        "% Despesas sem categoria": f"{uncategorized_rate * 100:.1f}%",
        "Tempo médio transação-aprovação": f"{avg_time:.1f} dias",
        "Ajustes manuais": f"{adjustments:.2f}",
    }


def render_table(title: str, records: Iterable[Dict[str, str]]) -> str:
    """Render a textual table representation used in CLI summary output."""

    lines = [title]
    for record in records:
        serialized = ", ".join(f"{key}: {value}" for key, value in record.items())
        lines.append(f"- {serialized}")
    return "\n".join(lines)


def render_reports(result: ReconciliationResult) -> str:
    """Render the reconciliation result in Markdown format (legacy compatibility)."""

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


__all__ = [
    "AggregatedFailure",
    "build_metadata",
    "build_reconciliation_result",
    "build_workbook",
    "render_reports",
    "render_summary",
    "render_table",
    "write_report",
]

