# Conciliador Bancário

Programa simples para comparar um extrato bancário com as movimentações do ERP.

## Uso

```bash
python conciliador.py extrato.csv movimentacoes_erp.csv
```

Os arquivos podem estar no formato CSV ou Excel (`.xlsx`). A conciliação é feita com base nas colunas em comum dos arquivos. Os registros que estiverem apenas em um dos arquivos são exibidos como diferenças.
