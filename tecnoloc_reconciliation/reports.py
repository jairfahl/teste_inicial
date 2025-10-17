main

from .models import ErpRecord, PayfyExpense, ReconciliationResult


main
    payload = {
        "Usuário": expense.user,
        "Data": expense.date.strftime("%d/%m/%Y %H:%M"),
        "Valor": f"{expense.value:.2f}",
        "Status": expense.status,
        "Categoria": expense.category,
        "ID": expense.expense_id or "",
        "Match": expense.match_type or "",
        "Motivo": (expense.failure_cause or expense.failure_reason or ""),
    }
main
    payload = {
        "Usuário": record.user,
        "Data": record.date.strftime("%d/%m/%Y %H:%M"),
        "Valor": f"{record.value:.2f}",
        "Tipo": record.erp_type,
        "Match": record.match_type or "",
        "Motivo": (record.failure_cause or record.failure_reason or ""),
    }
main
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


main
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
main
