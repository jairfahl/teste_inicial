"""Ferramenta simples de concilia\u00e7\u00e3o banc\u00e1ria.

Uso:
    python conciliador.py extrato.csv movimentacoes_erp.csv

Os arquivos podem estar em formato CSV ou Excel (.xlsx). A concilia\u00e7\u00e3o é feita com base nas colunas em comum entre os
arquivos e os registros que não estiverem presentes em ambos são exibidos
como diferenças.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd


def carregar_arquivo(caminho: str) -> pd.DataFrame:
    """Carrega um arquivo CSV ou Excel para um DataFrame.

    Parâmetros
    ----------
    caminho: str
        Caminho do arquivo a ser carregado.
    """
    arquivo = Path(caminho)
    if not arquivo.exists():
        raise FileNotFoundError(f"Arquivo n\u00e3o encontrado: {caminho}")

    if arquivo.suffix.lower() == ".csv":
        df = pd.read_csv(arquivo)
    else:
        df = pd.read_excel(arquivo)

    df.columns = [col.strip() for col in df.columns]
    return df


def conciliar(banco_df: pd.DataFrame, erp_df: pd.DataFrame) -> pd.DataFrame:
    """Retorna as diferenças entre dois DataFrames."""
    colunas_comuns = [c for c in banco_df.columns if c in erp_df.columns]
    if not colunas_comuns:
        raise ValueError("Os arquivos n\u00e3o possuem colunas em comum para concilia\u00e7\u00e3o.")

    conciliacao = banco_df.merge(erp_df, how="outer", indicator=True, on=colunas_comuns)
    diferencas = conciliacao[conciliacao["_merge"] != "both"]
    return diferencas


def main() -> None:
    if len(sys.argv) < 3:
        print("Uso: python conciliador.py extrato.csv movimentacoes_erp.csv")
        sys.exit(1)

    banco_df = carregar_arquivo(sys.argv[1])
    erp_df = carregar_arquivo(sys.argv[2])

    diferencas = conciliar(banco_df, erp_df)
    if diferencas.empty:
        print("Nenhuma diferen\u00e7a encontrada.")
    else:
        apenas_banco = diferencas[diferencas["_merge"] == "left_only"].drop(columns="_merge")
        apenas_erp = diferencas[diferencas["_merge"] == "right_only"].drop(columns="_merge")

        if not apenas_banco.empty:
            print("Presentes apenas no extrato banc\u00e1rio:")
            print(apenas_banco.to_string(index=False))

        if not apenas_erp.empty:
            print("Presentes apenas nas movimenta\u00e7\u00f5es do ERP:")
            print(apenas_erp.to_string(index=False))


if __name__ == "__main__":
    main()
