# 💵 Sistema de Orçamento (Excel/VBA + Python/MySQL)

Automação do **orçamento** das 4 lojas em duas frentes: **macros VBA** no Excel (organização/consolidação)
e um **ETL em Python** (Pandas + SQLAlchemy) que lê o `orcamento.xlsm` e carrega os dados em um **Data Warehouse MySQL**.

> ⚠️ **Aviso sobre os dados**
> Os dados neste repositório são **fictícios**, gerados apenas para demonstração. Os dados reais da
> operação em que o projeto foi usado são **confidenciais e estão protegidos** — nada real, credencial
> ou informação de terceiros foi incluído aqui.

## 🎯 Componentes
- `vba/Módulo*.bas` — macros de organização e consolidação orçamentária.
- `script_orcamento.py` — ETL que extrai datas/valores do Excel e faz **UPSERT** no banco (insere ou atualiza).
- `sql/schema.sql` — DDL do banco (dimensões + fato).

## 🏗️ Modelo dimensional (star schema)
```
   Dim_Loja      Dim_Mercadologico      Dim_Comprador      Dim_Tempo
      │                 │                     │                │
      └────────┬────────┴──────────┬──────────┴────────────────┘
               ▼                    ▼
                 Fato_OrcamentoCompras
     (meta × compra realizada · venda · CMV · % execução | grão: semana × loja × categoria × comprador)
```
- **Dimensões:** `Dim_Loja`, `Dim_Mercadologico`, `Dim_Comprador`, `Dim_Tempo` (grão semanal, 1..5 semanas/mês).
- **Fato:** `Fato_OrcamentoCompras` — meta vs. compra realizada, venda, CMV, diferença e % de execução da meta.
- Chave natural do UPSERT: `FK_Tempo + FK_Loja + FK_Mercadologico + FK_Comprador`.

DDL completo (chaves, FKs, índices e seed fictício de lojas) em [`sql/schema.sql`](sql/schema.sql).

## ▶️ Como rodar (Python)
```bash
pip install pandas numpy sqlalchemy pymysql openpyxl

# 1. cria o banco e as tabelas (o script já faz CREATE DATABASE 'orcamento')
mysql -u root < sql/schema.sql

# 2. executa o ETL (espera ./orcamento.xlsm; um exemplo fictício está em data_exemplo/)
python script_orcamento.py
```

## 📁 Estrutura
```
vba/Módulo*.bas          # macros de organização/consolidação no Excel
script_orcamento.py      # ETL Pandas + SQLAlchemy com UPSERT no MySQL
sql/schema.sql           # DDL do Data Warehouse (dimensões + fato)
data_exemplo/            # planilha de exemplo (dados fictícios)
```

---

### 🧰 Competências demonstradas
`Excel` · `VBA` · `Python` · `Pandas` · `SQLAlchemy` · `MySQL` · `ETL` · `Modelagem Dimensional`

### 👤 Autor
**José Vitor Santos Pinheiro** — Analista de Dados / BI / Ciência de Dados · vytorsantt@gmail.com
