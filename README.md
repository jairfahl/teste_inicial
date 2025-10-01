# Conciliação PayFy x ERP – Tecnoloc

Aplicativo de linha de comando para conciliação financeira automática entre PayFy e ERP Protheus, seguindo o fluxo descrito no prompt.

## Requisitos

- Python 3.11+

## Uso

1. Prepare os arquivos CSV exportados do PayFy e Protheus conforme estrutura indicada.
2. Execute:

```bash
python -m tecnoloc_reconciliation --period "01/01/2024 – 08:00" --cards cartoes.csv --expenses despesas.csv --erp erp.csv
```

O aplicativo imprimirá um relatório completo no terminal com resumos, correspondências e pendências.

## Estrutura

- `tecnoloc_reconciliation/loader.py`: leitura e validação dos arquivos.
- `tecnoloc_reconciliation/preprocess.py`: normalização, categorização e diagnósticos.
- `tecnoloc_reconciliation/matcher.py`: lógica de conciliação (exato, tolerância e agregações).
- `tecnoloc_reconciliation/reports.py`: geração dos relatórios finais.
- `tecnoloc_reconciliation/cli.py`: ponto de entrada da aplicação.

## Testes rápidos

Para executar um teste manual, utilize arquivos CSV fictícios e verifique o relatório impresso.
