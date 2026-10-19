# Sistema de Orçamento (Excel/VBA + Python/MySQL)

Automação do orçamento de compras das lojas em duas frentes: macros VBA no Excel, que organizam e
consolidam a planilha orçamentária, e um ETL em Python (Pandas + SQLAlchemy) que lê o `orcamento.xlsm`
e carrega os dados em um Data Warehouse MySQL modelado em star schema. O projeto conecta a ferramenta
que a operação já usa (Excel) a uma base analítica estruturada.

> **Nota de confidencialidade:** os dados presentes neste repositório são fictícios, gerados apenas
> para demonstração. Os dados reais da operação são confidenciais e estão protegidos — nenhum dado
> real, credencial ou informação de terceiros foi incluído aqui.

---

## Visão Geral

O orçamento de compras nasce em Excel e precisa virar informação analisável. Este projeto padroniza a
planilha com macros VBA e, em seguida, extrai datas e valores para carregar um modelo dimensional em
MySQL, com carga do tipo UPSERT (insere ou atualiza) para que reexecuções mantenham a base
consistente. O resultado é o acompanhamento de meta versus compra realizada com granularidade
semanal, por loja, categoria e comprador.

## Contexto de Negócio

O orçamento de compras é o que separa o estoque saudável do capital parado: comprar acima da meta
imobiliza dinheiro e aumenta o risco de perda; comprar abaixo gera ruptura. Controlar isso em
planilhas soltas dificulta comparar meta e realizado e enxergar quem está estourando o orçamento.
Estruturar o orçamento em um Data Warehouse devolve à área de compras a capacidade de acompanhar a
execução da meta em tempo de corrigir.

## O Problema que Resolve

- **Orçamento preso no Excel**, sem uma base analítica para comparar meta e realizado.
- **Consolidação manual** e sujeita a erro entre lojas e categorias.
- **Falta de granularidade** para identificar onde a meta de compra está sendo estourada.

## Público e Decisões Apoiadas

- **Compras:** acompanha a execução da meta por categoria e comprador e corrige desvios.
- **Diretoria:** enxerga o percentual de execução do orçamento por loja.
- **Controladoria:** compara meta, compra, venda e CMV a partir de uma base única.

## Impacto e Valor Gerado

- Conecta a planilha operacional (Excel) a um Data Warehouse consultável.
- Padroniza a consolidação orçamentária com macros VBA.
- Acompanha meta versus compra realizada com granularidade semanal por loja, categoria e comprador.
- Carga UPSERT: reexecuções atualizam sem duplicar.

---

## Arquitetura e Modelagem Dimensional (Star Schema)

```text
   Dim_Loja      Dim_Mercadologico      Dim_Comprador      Dim_Tempo
      |                 |                     |                |
      +--------+--------+----------+----------+----------------+
               v                   v
                 Fato_OrcamentoCompras
   (meta x compra realizada - venda - CMV - % execução | grão: semana x loja x categoria x comprador)
```

- **Dimensões:** `Dim_Loja`, `Dim_Mercadologico`, `Dim_Comprador` e `Dim_Tempo` (grão semanal, de 1 a
  5 semanas por mês).
- **Fato:** `Fato_OrcamentoCompras` — meta versus compra realizada, venda, CMV, diferença e percentual
  de execução da meta.
- **Chave natural do UPSERT:** `FK_Tempo + FK_Loja + FK_Mercadologico + FK_Comprador`.
- DDL completo (chaves, FKs, índices e seed fictício) em `sql/schema.sql`.

## Stack

Excel - VBA - Python - Pandas - SQLAlchemy - MySQL - ETL - Modelagem Dimensional.

## Como Rodar (Python)

```bash
pip install pandas numpy sqlalchemy pymysql openpyxl

# 1. cria o banco e as tabelas (o script já faz CREATE DATABASE 'orcamento')
mysql -u root < sql/schema.sql

# 2. executa o ETL (espera ./orcamento.xlsm; um exemplo fictício está em data_exemplo/)
python script_orcamento.py
```

## Estrutura do Projeto

```text
vba/Modulo*.bas       -> Macros de organização e consolidação no Excel
script_orcamento.py   -> ETL Pandas + SQLAlchemy com UPSERT no MySQL
sql/schema.sql        -> DDL do Data Warehouse (dimensões + fato)
data_exemplo/         -> Planilha de exemplo (dados fictícios)
```

## Autor

José Vitor Santos Pinheiro — Análise de Dados e Inteligência Comercial (Varejo, Indústria e Distribuição).
Contato: vytorsantt@gmail.com
