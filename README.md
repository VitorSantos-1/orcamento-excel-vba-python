# 💵 Sistema de Orçamento (Excel/VBA + Python/MySQL)

Automação do **orçamento** das 4 lojas em duas frentes: **macros VBA** no Excel (organização/consolidação)
e um **ETL em Python** (Pandas + SQLAlchemy) que lê o `orcamento.xlsm` e carrega os dados em um **banco MySQL**.

> ⚠️ **Aviso sobre os dados**
> Os dados neste repositório são **fictícios**, gerados apenas para demonstração. Os dados reais da
> operação em que o projeto foi usado são **confidenciais e estão protegidos** — nada real, credencial
> ou informação de terceiros foi incluído aqui.

## 🎯 Componentes
- `vba/Módulo*.bas` — macros de organização e consolidação orçamentária.
- `script_orcamento.py` — ETL que extrai as datas/valores do Excel e grava no MySQL.

## ▶️ Como rodar (Python)
```bash
pip install pandas sqlalchemy pymysql openpyxl
python script_orcamento.py     # espera ./orcamento.xlsm e um banco 'orcamento'
```

---

### 🧰 Competências demonstradas
`Excel` · `VBA` · `Python` · `Pandas` · `SQLAlchemy` · `MySQL` · `ETL`

### 👤 Autor
**José Vitor Santos Pinheiro** — Analista de Dados / BI / Ciência de Dados · vytorsantt@gmail.com
