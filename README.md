# Conciliação PayFy x ERP – Tecnoloc

Aplicativo de linha de comando para conciliação financeira automática entre PayFy e ERP Protheus, seguindo o fluxo descrito no prompt.

## Requisitos

- Python 3.11+

## Uso

1. Garanta que as planilhas XLSX exportadas estejam no diretório `/downloads/Codex/TLC` com os nomes:
   - `cartao tecnicos.xlsx`
   - `despesas.xlsx`
   - `saldo empresa.xlsx`
2. Execute:

```bash
python -m tecnoloc_reconciliation --period "01/01/2024 – 08:00" \
  --out "/downloads/Codex/TLC/relatorio_conciliacao.xlsx"
```

O aplicativo utilizará automaticamente os caminhos padrão acima (você pode sobrescrevê-los com `--cards`, `--expenses` e `--erp` se necessário), imprimirá um relatório completo no terminal e, quando `--out` for fornecido, exportará também um arquivo Excel multiaba com os resultados.

## Estrutura

- `tecnoloc_reconciliation/loader.py`: leitura e validação dos arquivos.
- `tecnoloc_reconciliation/preprocess.py`: normalização, categorização e diagnósticos.
- `tecnoloc_reconciliation/matcher.py`: lógica de conciliação (exato, tolerância e agregações).
- `tecnoloc_reconciliation/reports.py`: geração dos relatórios finais.
- `tecnoloc_reconciliation/cli.py`: ponto de entrada da aplicação.

## Testes rápidos

Para executar um teste manual, utilize arquivos XLSX fictícios e verifique o relatório impresso.
