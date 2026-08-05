# Importando as bibliotecas
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import pathlib
from sqlalchemy import create_engine, text
import calendar

# ------------------------------------------------------------------------------
# 1. AJUSTANDO O BANCO

# Aqui é onde colocamos o usuario [depois do root:] e a senha [onde fica o 3306] do banco
DB_CONNECTION = 'mysql+pymysql://root:@localhost:3306/orcamento'
engine = create_engine(DB_CONNECTION)
BASE_DIR = pathlib.Path().resolve()
INPUT_FILE = BASE_DIR / "orcamento.xlsm"

# ------------------------------------------------------------------------------
# 2. FUNÇÃO: LER DATAS DA ABA 'LOJA CENTRAL'
def extrair_datas_reais_do_excel(arquivo):
    print("🔎 Lendo datas reais da aba 'loja central'...")
    try:
        df = pd.read_excel(arquivo, sheet_name='loja central', header=None)
        
        # Mapeia: {Numero_Semana: (Indice_Linha_DE, Indice_Linha_PARA)}
        indices = {
            1: (0, 1),   # Linhas 1 e 2 do Excel
            2: (4, 5),   # Linhas 5 e 6 do Excel
            3: (8, 9),   # Linhas 9 e 10 do Excel
            4: (12, 13), # Linhas 13 e 14 do Excel
            5: (16, 17)  # Linhas 17 e 18 do Excel
        }
        
        mapa_datas = {}
        for num_sem, (idx_de, idx_para) in indices.items():
            if idx_para < len(df):
                val_de = df.iloc[idx_de, 1] 
                val_para = df.iloc[idx_para, 1]
                
                if isinstance(val_de, (datetime, pd.Timestamp)) and isinstance(val_para, (datetime, pd.Timestamp)):
                    # Agora guardamos o início e o fim lidos direto da planilha
                    mapa_datas[num_sem] = {'inicio': val_de, 'fim': val_para}
                    print(f"    Semana {num_sem}: {val_de.strftime('%d/%m/%Y')} até {val_para.strftime('%d/%m/%Y')}")
                else:
                    print(f"    Semana {num_sem}: Data inválida ou vazia.")
        return mapa_datas
    except Exception as e:
        print(f" Erro ao ler datas: {e}")
        return {}

# ------------------------------------------------------------------------------
# 3. BANCO DE DADOS DIMENSÕES E TEMPO

def get_or_create_dimension(engine, table_name, key_col, val_col, df_source):
    """Sincroniza dimensões (Loja, Produto, Comprador)"""
    if df_source.empty: return {}
    
    # 1. Busca existentes
    with engine.connect() as conn:
        try:
            df_db = pd.read_sql(f"SELECT {key_col}, {val_col} FROM {table_name}", conn)
        except:
            df_db = pd.DataFrame(columns=[key_col, val_col])
    
    existing_map = dict(zip(df_db[val_col].astype(str), df_db[key_col]))
    
    # 2. Identifica novos
    df_source[val_col] = df_source[val_col].astype(str)
    novos = df_source[~df_source[val_col].isin(existing_map.keys())].drop_duplicates(subset=[val_col])
    
    # 3. Insere novos
    if not novos.empty:
        print(f"   [Dimensão] Inserindo {len(novos)} novos registros em {table_name}...")
        # Prepara colunas para inserção (ajuste conforme seu banco se tiver ID obrigatório)
        cols_to_insert = [c for c in novos.columns if c != key_col] # Remove PK se existir
        novos_limpo = novos[cols_to_insert].copy()
        
        with engine.begin() as conn:
            novos_limpo.to_sql(table_name, conn, if_exists='append', index=False)
        
        # Recarrega mapa
        with engine.connect() as conn:
            df_db = pd.read_sql(f"SELECT {key_col}, {val_col} FROM {table_name}", conn)
        existing_map = dict(zip(df_db[val_col].astype(str), df_db[key_col]))
        
    return existing_map

def get_tempo_map_dinamico(engine, mapa_semanas_excel):
    """Garante que as datas do Excel existem na Dim_Tempo do banco usando as datas exatas da planilha"""
    # Extrai apenas as datas de início para a lista de verificação
    dates_list = [dados['inicio'] for dados in mapa_semanas_excel.values()]
    dates_str = [d.strftime('%Y-%m-%d') for d in dates_list]
    
    query = "SELECT PK_Tempo, Data_Inicio_Semana FROM Dim_Tempo"
    with engine.connect() as conn:
        df_tempo = pd.read_sql(query, conn)
    
    df_tempo['Data_Str'] = pd.to_datetime(df_tempo['Data_Inicio_Semana']).dt.strftime('%Y-%m-%d')
    tempo_map = dict(zip(df_tempo['Data_Str'], df_tempo['PK_Tempo']))
    
    missing = [d for d in dates_str if d not in tempo_map]
    
    if missing:
        print(f"   [Tempo] Criando {len(missing)} novas semanas na Dim_Tempo...")
        new_rows = []
        
        for num_sem, dados in mapa_semanas_excel.items():
            dt_inicio = dados['inicio']
            dt_fim = dados['fim']
            d_str = dt_inicio.strftime('%Y-%m-%d')
            
            if d_str in missing:
                new_rows.append({
                    'Data_Inicio_Semana': dt_inicio,
                    'Data_Fim_Semana': dt_fim,  # <--- USANDO A DATA 'PARA' EXATA DA PLANILHA
                    'Ano': dt_inicio.year,
                    'Mes': dt_inicio.month,
                    'Nome_Mes': dt_inicio.strftime('%B'),
                    'AnoMes': dt_inicio.strftime('%Y%m'),
                    'Numero_Semana': num_sem,
                    'Descricao_Semana': 'Semana Importada'
                })
        with engine.begin() as conn:
            pd.DataFrame(new_rows).to_sql('Dim_Tempo', conn, if_exists='append', index=False)
        
        # Atualiza mapa
        with engine.connect() as conn:
            df_tempo = pd.read_sql(query, conn)
        df_tempo['Data_Str'] = pd.to_datetime(df_tempo['Data_Inicio_Semana']).dt.strftime('%Y-%m-%d')
        tempo_map = dict(zip(df_tempo['Data_Str'], df_tempo['PK_Tempo']))
    return tempo_map

# ------------------------------------------------------------------------------
# 4. EXECUÇÃO PRINCIPAL

print(" Iniciando Processo de Carga...")

# A. Ler Datas Reais
datas_reais = extrair_datas_reais_do_excel(INPUT_FILE)
if not datas_reais:
    print("❌ Erro critico: Datas não encontradas.")
    exit()

# B. Ler Planilhas de Orçamento
print("\n📂 Lendo planilhas de orçamento...")
abas = {
    'loja central': pd.read_excel(INPUT_FILE, sheet_name='orc_loja central'),
    'loja shopping': pd.read_excel(INPUT_FILE, sheet_name='orc_loja shopping'),
    'loja sul': pd.read_excel(INPUT_FILE, sheet_name='orc_loja sul'),
    'loja oeste': pd.read_excel(INPUT_FILE, sheet_name='orc_loja oeste')
}

# C. Sincronizar Dimensões
print("\n Atualizando Dimensões...")

# Loja
df_loja = pd.DataFrame([
    {'Codigo_Loja': 1, 'Nome_Loja': 'Loja Central'},
    {'Codigo_Loja': 2, 'Nome_Loja': 'Loja Shopping'},
    {'Codigo_Loja': 3, 'Nome_Loja': 'Loja Sul'},
    {'Codigo_Loja': 4, 'Nome_Loja': 'Loja Oeste'}
])
map_loja = get_or_create_dimension(engine, 'Dim_Loja', 'PK_Loja', 'Nome_Loja', df_loja)

# Mercadologico (Junta tudo para criar a dimensão completa)
dfs_merc = []
for sufixo, df in abas.items():
    if not df.empty:
        temp = df[[f'id_mercadologico_{sufixo}', f'mercadologico_{sufixo}']].copy()
        temp.columns = ['Id_Mercadologico', 'Nome_Mercadologico']
        dfs_merc.append(temp)
df_merc_all = pd.concat(dfs_merc).dropna().drop_duplicates()
map_merc = get_or_create_dimension(engine, 'Dim_Mercadologico', 'PK_Mercadologico', 'Nome_Mercadologico', df_merc_all)

# Comprador
dfs_comp = []
for df in abas.values():
    if not df.empty:
        temp = df[['id_comprador', 'comprador']].copy()
        temp.columns = ['Id_Comprador', 'Nome_Comprador']
        dfs_comp.append(temp)
df_comp_all = pd.concat(dfs_comp).dropna().drop_duplicates()
map_comp = get_or_create_dimension(engine, 'Dim_Comprador', 'PK_Comprador', 'Nome_Comprador', df_comp_all)

# Tempo (Baseado nas datas do Excel)
map_tempo = get_tempo_map_dinamico(engine, datas_reais)

# D. Processar Fatos
print("\n📊 Processando dados e preparando carga...")
fatos = []

for sufixo, df in abas.items():
    # Mapeia nome da aba para nome da loja
    nome_loja_map = {
        'loja central': 'Loja Central', 'loja shopping': 'Loja Shopping', 
        'loja sul': 'Loja Sul', 'loja oeste': 'Loja Oeste'
    }
    pk_loja = map_loja.get(nome_loja_map[sufixo])
    if not pk_loja: continue

    for _, row in df.iterrows():
        # Busca IDs
        nome_prod = str(row.get(f'mercadologico_{sufixo}'))
        nome_comp = str(row.get('comprador'))
        pk_prod = map_merc.get(nome_prod)
        pk_comp_id = map_comp.get(nome_comp)        
        if not pk_prod or not pk_comp_id: continue

        # Loop semanas 1 a 5
        for i in range(1, 6):
            if i not in datas_reais: continue
            
            # 1. PEGA OS VALORES
            raw_meta = row.get(f'meta_orcamento_compras_{i}semana_{sufixo}')
            raw_compra = row.get(f'{i}semana_compra_{sufixo}')
            venda_sem = row.get(f'venda_{i}semana_{sufixo}')

            # =================================================================
            # NOVA LÓGICA: IGNORAR SEMANAS VAZIAS! 
            # Se a meta é nula, a compra é nula e a venda é zero, a gente PULA.
            # É isso que evita que ele crie o mês de Março atoa.
            # =================================================================
            if pd.isnull(raw_meta) and pd.isnull(raw_compra) and (pd.isnull(venda_sem) or venda_sem == 0):
                continue
            data_str = datas_reais[i]['inicio'].strftime('%Y-%m-%d')
            fk_tempo = map_tempo.get(data_str)
            if not fk_tempo: continue

            # Faz os cálculos
            if pd.notnull(raw_compra):
                val_meta = float(raw_meta) if pd.notnull(raw_meta) else 0.0
                val_compra = float(raw_compra)
                diferenca = val_meta - val_compra
                if val_meta != 0:
                    percentual = val_compra / val_meta
                else:
                    percentual = 0.0
            else:
                diferenca = None

                percentual = None
            # 2. MONTA A LINHA
            fatos.append({
                'FK_Tempo': fk_tempo,
                'FK_Loja': pk_loja,
                'FK_Mercadologico': pk_prod,
                'FK_Comprador': pk_comp_id,
                'Tipo_Semana': i,
                'Margem_Percentual': row.get(f'margem_{sufixo}'),
                'Complemento_Margem': row.get(f'compl/margem_{sufixo}'),
                'Orcamento_Compras_Total': row.get(f'orcamento_de_compras_{sufixo}'),
                'Total_CMV_Periodo': row.get(f'total_cmv_{sufixo}'),
                'Total_Orcamento_Periodo': row.get(f'total_orcamento_{sufixo}'),
                'Total_Compra_Semanal': row.get(f'total_compra_semanal_{sufixo}'),
                'Venda_Semana': venda_sem,
                'CMV_Semana': row.get(f'cmv_{i}semana_{sufixo}'),
                'Part_CMV_Semana': row.get(f'part_cmv_{i}semana_{sufixo}'),
                'Meta_Orcamento_Semana': raw_meta,
                'Compra_Realizada_Semana': raw_compra,
                'Diferenca_Meta_Compra': diferenca,
                'Percentual_Execucao_Meta': percentual
            })

df_final = pd.DataFrame(fatos)
df_final = df_final.replace({np.nan: None})
# E. UPSERT (INSERIR OU ATUALIZAR)
print(f"\n🔄 Sincronizando {len(df_final)} registros no Banco...")

if not df_final.empty:
    # Busca IDs de fatos existentes para esse período
    tempos_ids = tuple(set(df_final['FK_Tempo']))
    if len(tempos_ids) == 1: tempos_ids = f"({tempos_ids[0]})"
    else: tempos_ids = str(tempos_ids)
    
    query_check = f"""
        SELECT PK_Fato, FK_Tempo, FK_Loja, FK_Mercadologico, FK_Comprador 
        FROM Fato_OrcamentoCompras 
        WHERE FK_Tempo IN {tempos_ids}
    """
    with engine.connect() as conn:
        existing = pd.read_sql(query_check, conn)
    
    # Merge para achar inserts vs updates
    merged = pd.merge(df_final, existing, on=['FK_Tempo', 'FK_Loja', 'FK_Mercadologico', 'FK_Comprador'], how='left', indicator=True)
    
    to_insert = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge', 'PK_Fato'])
    to_update = merged[merged['_merge'] == 'both']
    
    with engine.begin() as conn:
        # Insert (Novos registros)
        if not to_insert.empty:
            print(f"   ➕ Inserindo {len(to_insert)} novos registros...")
            to_insert.to_sql('Fato_OrcamentoCompras', conn, if_exists='append', index=False)
        
        # Update (Atualiza valores que mudaram) - DEIXE APENAS ESTE BLOCO AQUI EMBAIXO 👇
        if not to_update.empty:
            print(f"   🔄 Atualizando {len(to_update)} registros existentes...")
            for _, row in to_update.iterrows():
                # Atualiza campos vitais + os 2 novos calculados
                sql = text("""
                    UPDATE Fato_OrcamentoCompras SET 
                        Venda_Semana = :venda,
                        CMV_Semana = :cmv,
                        Meta_Orcamento_Semana = :meta,
                        Compra_Realizada_Semana = :compra,
                        Diferenca_Meta_Compra = :diferenca,
                        Percentual_Execucao_Meta = :percentual
                    WHERE PK_Fato = :pk
                """)
                conn.execute(sql, {
                    'venda': row['Venda_Semana'], 
                    'cmv': row['CMV_Semana'],
                    'meta': row['Meta_Orcamento_Semana'], 
                    'compra': row['Compra_Realizada_Semana'],
                    'diferenca': row['Diferenca_Meta_Compra'],    
                    'percentual': row['Percentual_Execucao_Meta'], 
                    'pk': row['PK_Fato']
                })
print("\n✅ Processo finalizado com sucesso!")