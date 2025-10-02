"""Command line interface for the Tecnoloc reconciliation app."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from . import loader, matcher, preprocess, reports
from .models import ErpRecord, PayfyExpense


DEFAULT_DATA_DIR = Path("/downloads/Codex/TLC")
DEFAULT_CARD_FILE = DEFAULT_DATA_DIR / "cartao tecnicos.xlsx"
DEFAULT_EXPENSE_FILE = DEFAULT_DATA_DIR / "despesas.xlsx"
DEFAULT_ERP_FILE = DEFAULT_DATA_DIR / "saldo empresa.xlsx"


def _parse_period(value: str | None) -> datetime:
    if not value:
        raise preprocess.PeriodNotProvidedError("Período de conciliação não informado.")
    try:
        return datetime.strptime(value, "%d/%m/%Y – %H:%M")
    except ValueError as exc:
        raise preprocess.PeriodNotProvidedError(
            "Período deve estar no formato 'dd/mm/aaaa – hh:mm'."
        ) from exc


def _prepare_data(
    period: datetime,
    card_path: Path,
    expenses_path: Path,
    erp_path: Path,
) -> tuple[List[PayfyExpense], List[ErpRecord]]:
    # Carrega todos os arquivos para garantir a validação do conteúdo.
    loader.load_payfy_card_summary(card_path)
    payfy_expenses = loader.load_payfy_expenses(expenses_path)
    erp_records = loader.load_erp_records(erp_path)

    preprocess.normalize_entry_types(payfy_expenses)
    preprocess.normalize_entry_types(erp_records)
    preprocess.map_categories(payfy_expenses)
    preprocess.apply_date_rules(payfy_expenses)
    preprocess.ensure_status_validated(payfy_expenses)
    preprocess.validate_period(payfy_expenses, period)
    preprocess.detect_duplicates(payfy_expenses)

    return payfy_expenses, erp_records


def run_app(args: argparse.Namespace) -> tuple[str, Optional[Path]]:
    period = _parse_period(args.period)
    card_path = Path(args.cards)
    expenses_path = Path(args.expenses)
    erp_path = Path(args.erp)

    if not card_path.exists() or not card_path.is_file():
        raise FileNotFoundError(f"Arquivo de cartões não encontrado: {card_path}")
    if not expenses_path.exists() or not expenses_path.is_file():
        raise FileNotFoundError(f"Arquivo de despesas não encontrado: {expenses_path}")
    if not erp_path.exists() or not erp_path.is_file():
        raise FileNotFoundError(f"Arquivo ERP não encontrado: {erp_path}")

    payfy_expenses, erp_records = _prepare_data(period, card_path, expenses_path, erp_path)
    matcher.reconcile(payfy_expenses, erp_records)
    diagnostics = preprocess.summarize_failures(payfy_expenses, erp_records)
    result = reports.build_reconciliation_result(payfy_expenses, erp_records, diagnostics)
    export_path: Optional[Path] = None
    if getattr(args, "out", None):
        payload = reports.build_export_payload(result)
        export_path = reports.export_excel(payload, args.out)
    return reports.render_reports(result), export_path


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Conciliação PayFy x ERP – Tecnoloc")
    parser.add_argument("--period", help="Período de conciliação (dd/mm/aaaa – hh:mm)")
    parser.add_argument(
        "--cards",
        default=str(DEFAULT_CARD_FILE),
        help=f"Planilha de cartão técnicos (XLSX). Padrão: {DEFAULT_CARD_FILE}",
    )
    parser.add_argument(
        "--expenses",
        default=str(DEFAULT_EXPENSE_FILE),
        help=f"Planilha de despesas PayFy (XLSX). Padrão: {DEFAULT_EXPENSE_FILE}",
    )
    parser.add_argument(
        "--erp",
        default=str(DEFAULT_ERP_FILE),
        help=f"Planilha de saldo ERP (XLSX). Padrão: {DEFAULT_ERP_FILE}",
    )
    parser.add_argument(
        "--out",
        help="Arquivo Excel de saída (XLSX) para exportar o relatório consolidado.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_argument_parser()
    args = parser.parse_args(argv)
    try:
        report, export_path = run_app(args)
        print(report)
        if export_path:
            print(f"Arquivo Excel gerado em: {export_path}")
        return 0
    except preprocess.PeriodNotProvidedError as error:
        parser.error(str(error))
    except loader.DataValidationError as error:
        parser.error(str(error))
    except FileNotFoundError as error:
        parser.error(str(error))

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
