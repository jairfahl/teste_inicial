# Conciliação PayFy x ERP – Tecnoloc

Aplicativo de linha de comando para conciliação financeira automática entre PayFy e ERP Protheus, seguindo o fluxo descrito no prompt.

## Requisitos

- Python 3.11+

## Uso

main

## Estrutura

- `tecnoloc_reconciliation/loader.py`: leitura e validação dos arquivos.
- `tecnoloc_reconciliation/preprocess.py`: normalização, categorização e diagnósticos.
- `tecnoloc_reconciliation/matcher.py`: lógica de conciliação (exato, tolerância e agregações).
- `tecnoloc_reconciliation/reports.py`: geração dos relatórios finais.
- `tecnoloc_reconciliation/cli.py`: ponto de entrada da aplicação.

## Testes rápidos

 main
