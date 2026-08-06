-- ============================================================================
--  DATA WAREHOUSE — ORÇAMENTO DE COMPRAS (4 lojas)
--  Modelo dimensional (star schema) — MySQL 8+ / InnoDB
--
--  Reconstruído por engenharia reversa a partir de script_orcamento.py
--  (ETL Pandas + SQLAlchemy que lê o orcamento.xlsm e carrega este banco).
--
--  As chaves primárias são AUTO_INCREMENT porque o ETL usa a estratégia
--  get_or_create_dimension (insere o nome e deixa o banco gerar o PK).
-- ============================================================================

CREATE DATABASE IF NOT EXISTS orcamento
  CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE orcamento;

-- ----------------------------------------------------------------------------
--  DIMENSÃO: LOJA  (Central, Shopping, Sul, Oeste)
-- ----------------------------------------------------------------------------
CREATE TABLE Dim_Loja (
  PK_Loja      INT          NOT NULL AUTO_INCREMENT,
  Codigo_Loja  INT          NULL,
  Nome_Loja    VARCHAR(60)  NOT NULL,
  PRIMARY KEY (PK_Loja),
  UNIQUE KEY uq_loja_nome (Nome_Loja)
) ENGINE=InnoDB;

-- ----------------------------------------------------------------------------
--  DIMENSÃO: MERCADOLÓGICO (categoria)
-- ----------------------------------------------------------------------------
CREATE TABLE Dim_Mercadologico (
  PK_Mercadologico    INT           NOT NULL AUTO_INCREMENT,
  Id_Mercadologico    VARCHAR(20)   NULL,
  Nome_Mercadologico  VARCHAR(120)  NOT NULL,
  PRIMARY KEY (PK_Mercadologico),
  UNIQUE KEY uq_merc_nome (Nome_Mercadologico)
) ENGINE=InnoDB;

-- ----------------------------------------------------------------------------
--  DIMENSÃO: COMPRADOR
-- ----------------------------------------------------------------------------
CREATE TABLE Dim_Comprador (
  PK_Comprador    INT          NOT NULL AUTO_INCREMENT,
  Id_Comprador    VARCHAR(20)  NULL,
  Nome_Comprador  VARCHAR(80)  NOT NULL,
  PRIMARY KEY (PK_Comprador),
  UNIQUE KEY uq_comp_nome (Nome_Comprador)
) ENGINE=InnoDB;

-- ----------------------------------------------------------------------------
--  DIMENSÃO: TEMPO  (grão semanal — 5 semanas por mês/loja)
-- ----------------------------------------------------------------------------
CREATE TABLE Dim_Tempo (
  PK_Tempo            INT          NOT NULL AUTO_INCREMENT,
  Data_Inicio_Semana  DATE         NOT NULL,
  Data_Fim_Semana     DATE         NULL,
  Ano                 SMALLINT     NULL,
  Mes                 TINYINT      NULL,
  Nome_Mes            VARCHAR(20)  NULL,
  AnoMes              CHAR(6)      NULL,   -- 'AAAAMM'
  Numero_Semana       TINYINT      NULL,   -- 1 a 5
  Descricao_Semana    VARCHAR(60)  NULL,
  PRIMARY KEY (PK_Tempo),
  UNIQUE KEY uq_tempo_inicio (Data_Inicio_Semana)
) ENGINE=InnoDB;

-- ----------------------------------------------------------------------------
--  FATO: ORÇAMENTO DE COMPRAS
--  Grão: tempo(semana) × loja × mercadológico × comprador.
--  Chave natural (usada no UPSERT do ETL): FK_Tempo + FK_Loja +
--  FK_Mercadologico + FK_Comprador.
-- ----------------------------------------------------------------------------
CREATE TABLE Fato_OrcamentoCompras (
  PK_Fato                   BIGINT         NOT NULL AUTO_INCREMENT,
  FK_Tempo                  INT            NOT NULL,
  FK_Loja                   INT            NOT NULL,
  FK_Mercadologico          INT            NOT NULL,
  FK_Comprador              INT            NOT NULL,
  Tipo_Semana               TINYINT        NULL,          -- número da semana (1..5)
  -- Indicadores de período
  Margem_Percentual         DECIMAL(9,4)   NULL,
  Complemento_Margem        DECIMAL(9,4)   NULL,
  Orcamento_Compras_Total   DECIMAL(16,2)  NULL,
  Total_CMV_Periodo         DECIMAL(16,2)  NULL,
  Total_Orcamento_Periodo   DECIMAL(16,2)  NULL,
  Total_Compra_Semanal      DECIMAL(16,2)  NULL,
  -- Medidas da semana
  Venda_Semana              DECIMAL(16,2)  NULL,
  CMV_Semana                DECIMAL(16,2)  NULL,
  Part_CMV_Semana           DECIMAL(9,4)   NULL,
  Meta_Orcamento_Semana     DECIMAL(16,2)  NULL,
  Compra_Realizada_Semana   DECIMAL(16,2)  NULL,
  -- Cálculos derivados (meta x realizado)
  Diferenca_Meta_Compra     DECIMAL(16,2)  NULL,
  Percentual_Execucao_Meta  DECIMAL(9,4)   NULL,
  PRIMARY KEY (PK_Fato),
  UNIQUE KEY uq_fato_orcamento (FK_Tempo, FK_Loja, FK_Mercadologico, FK_Comprador),
  CONSTRAINT fk_fato_tempo    FOREIGN KEY (FK_Tempo)        REFERENCES Dim_Tempo (PK_Tempo),
  CONSTRAINT fk_fato_loja     FOREIGN KEY (FK_Loja)         REFERENCES Dim_Loja (PK_Loja),
  CONSTRAINT fk_fato_merc     FOREIGN KEY (FK_Mercadologico) REFERENCES Dim_Mercadologico (PK_Mercadologico),
  CONSTRAINT fk_fato_comp     FOREIGN KEY (FK_Comprador)    REFERENCES Dim_Comprador (PK_Comprador)
) ENGINE=InnoDB;

CREATE INDEX ix_fato_tempo ON Fato_OrcamentoCompras (FK_Tempo);
CREATE INDEX ix_fato_loja  ON Fato_OrcamentoCompras (FK_Loja);

-- ============================================================================
--  SEED — lojas fixas do projeto (dados fictícios, conforme aviso do README)
-- ============================================================================
INSERT INTO Dim_Loja (Codigo_Loja, Nome_Loja) VALUES
  (1, 'Loja Central'),
  (2, 'Loja Shopping'),
  (3, 'Loja Sul'),
  (4, 'Loja Oeste');
