"""Command line interface for the Tecnoloc reconciliation app."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

from . import loader, matcher, preprocess, reports
from .models import ErpRecord, PayfyExpense


def _prepare_data(
    card_path: Path,
    expenses_path: Path,
    erp_path: Path,
) -> Tuple[List[PayfyExpense], List[ErpRecord]]:
    """Load and preprocess the PayFy and ERP data."""

    # Carrega todos os arquivos para garantir a validação do conteúdo.
    loader.load_payfy_card_summary(card_path)
    payfy_expenses = loader.load_payfy_expenses(expenses_path)
    erp_records = loader.load_erp_records(erp_path)

    preprocess.normalize_entry_types(payfy_expenses)
    preprocess.normalize_entry_types(erp_records)
    preprocess.map_categories(payfy_expenses)
    preprocess.apply_date_rules(payfy_expenses)
    preprocess.ensure_status_validated(payfy_expenses)
    preprocess.detect_duplicates(payfy_expenses)

    return payfy_expenses, erp_records


def _month_year_key(date: datetime) -> Tuple[int, int]:
    return date.year, date.month


def _extract_payfy_competence(expenses: Iterable[PayfyExpense]) -> Sequence[Tuple[int, int]]:
    months = {
        _month_year_key(expense.approval_date)
        for expense in expenses
        if expense.approval_date is not None
    }
    if not months:
        raise preprocess.PeriodNotProvidedError(
            "Não foi possível identificar a competência pelas aprovações do PayFy."
        )
    return sorted(months)


def _extract_protheus_competence(movements: Sequence[datetime], erp_records: Iterable[ErpRecord]) -> Sequence[Tuple[int, int]]:
    months = {_month_year_key(moment) for moment in movements}
    if not months:
        raw_months = {
            _month_year_key(record.date)
            for record in erp_records
        }
        months.update(raw_months)
    if not months:
        raise preprocess.PeriodNotProvidedError(
            "Não foi possível identificar a competência pelos lançamentos do Protheus."
        )
    return sorted(months)


def _confirm_competence(possible_months: Sequence[Tuple[int, int]]) -> datetime:
    options = ", ".join(f"{month:02d}/{year}" for year, month in possible_months)
    print(f"Competências detectadas: {options}.")
    while True:
        response = input("Confirmar MM/AAAA: ").strip()
        if not response:
            print("Informe a competência no formato MM/AAAA.")
            continue
        try:
            confirmed = datetime.strptime(response, "%m/%Y")
        except ValueError:
            print("Formato inválido. Utilize MM/AAAA.")
            continue
        return confirmed.replace(day=1, hour=0, minute=0, second=0, microsecond=0)


def _determine_competence(
    payfy_expenses: Sequence[PayfyExpense],
    erp_records: Sequence[ErpRecord],
    protheus_movements: Sequence[datetime],
) -> datetime:
    payfy_months = _extract_payfy_competence(payfy_expenses)
    protheus_months = _extract_protheus_competence(protheus_movements, erp_records)

    combined = sorted(set(payfy_months) | set(protheus_months))
    if not combined:
        raise preprocess.PeriodNotProvidedError(
            "Não foi possível determinar a competência de conciliação."
        )
    if len(combined) == 1:
        year, month = combined[0]
        return datetime(year, month, 1)

    return _confirm_competence(combined)


def _ensure_input_file(path: Path, *, description: str, parser: argparse.ArgumentParser) -> Path:
    if not path.exists() or not path.is_file():
        parser.error(f"Arquivo {description} não encontrado: {path}")
    return path


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Conciliação PayFy x ERP – Tecnoloc")
    parser.add_argument(
        "--payfy-card-summary",
        required=True,
        type=Path,
        help="Caminho para o resumo de cartões PayFy.",
    )
    parser.add_argument(
        "--payfy-expenses",
        required=True,
        type=Path,
        help="Caminho para a planilha de despesas detalhadas do PayFy.",
    )
    parser.add_argument(
        "--erp-records",
        required=True,
        type=Path,
        help="Caminho para o relatório financeiro do Protheus.",
    )
    parser.add_argument(
        "--erp-movements",
        required=True,
        type=Path,
        help="Caminho para a planilha de movimentos do Protheus (campo data_mov).",
    )
    parser.add_argument(
        "--output",
        required=True,
        type=Path,
        help="Caminho do arquivo Excel a ser gerado com o resultado da conciliação.",
    )
    return parser


def main(argv: List[str] | None = None) -> int:
    parser = build_argument_parser()
    args = parser.parse_args(argv)

    try:
        card_path = _ensure_input_file(args.payfy_card_summary, description="de cartões do PayFy", parser=parser)
        expenses_path = _ensure_input_file(args.payfy_expenses, description="de despesas do PayFy", parser=parser)
        erp_path = _ensure_input_file(args.erp_records, description="financeiro do Protheus", parser=parser)
        movements_path = _ensure_input_file(args.erp_movements, description="de movimentos do Protheus", parser=parser)

        payfy_expenses, erp_records = _prepare_data(card_path, expenses_path, erp_path)
        protheus_movements = loader.load_erp_movements(movements_path)

        competence = _determine_competence(payfy_expenses, erp_records, protheus_movements)
        preprocess.validate_period(payfy_expenses, competence)

        matcher.reconcile(payfy_expenses, erp_records)
        diagnostics = preprocess.summarize_failures(payfy_expenses, erp_records)
        result = reports.build_reconciliation_result(payfy_expenses, erp_records, diagnostics)

        output_path: Path = args.output
        output_path.parent.mkdir(parents=True, exist_ok=True)
        reports.write_excel_report(result, output_path)
        return 0
    except preprocess.PeriodNotProvidedError as error:
        parser.error(str(error))
    except loader.DataValidationError as error:
        parser.error(str(error))
    except FileNotFoundError as error:
        parser.error(str(error))
    except RuntimeError as error:
        parser.error(str(error))
    except OSError as error:
        parser.error(f"Falha ao gravar o relatório: {error}")

    return 1


if __name__ == "__main__":
    raise SystemExit(main())

