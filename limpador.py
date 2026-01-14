import pandas as pd
import os
import re

# --- CONFIGURAÇÕES ---
EXCEL_PATH = 'resultado_measure_killer.xlsx'
PBIP_TABLES_PATH = 'Panorama Geral e OS Abertas.SemanticModel/definition/tables'

# Tabelas que o script NUNCA deve tocar, por sensibilidade de dados que podem quebrar o BI
TABELAS_FIXAS_EXCECAO = ['d_calendario']

def clean_tmdl():
    # 1. Carregar o Excel
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        print(f"-- Erro ao ler o arquivo Excel: {e}")
        return

    # --- LÓGICA DE DETECÇÃO DE COLUNAS ---
    # O Measure Killer muda esses nomes dependendo da versão/tipo de exportação, essa lista serve pra listar os possíveis nomes das colunas
    coluna_status = next((c for c in ['is_used', 'Status', 'is used', 'status'] if c in df.columns), None)
    coluna_tipo = next((c for c in ['type', 'Type', 'object_type', 'Category'] if c in df.columns), None)
    coluna_tabela = next((c for c in ['table', 'Table', 'table_name', 'Table Name'] if c in df.columns), None)
    coluna_nome = next((c for c in ['name', 'Name', 'object_name'] if c in df.columns), None)

    if not all([coluna_status, coluna_tipo, coluna_tabela, coluna_nome]):
        print(f"-- Erro: Colunas essenciais não encontradas.")
        print(f"-- Colunas detectadas no seu Excel: {df.columns.tolist()}")
        return
    
    print(f"- Mapeamento detectado: Status='{coluna_status}', Tipo='{coluna_tipo}', Tabela='{coluna_tabela}'")

    # 2. Filtragem de Sujeira
    df[coluna_status] = df[coluna_status].astype(str).str.strip()
    df[coluna_tipo] = df[coluna_tipo].astype(str).str.strip()
    
    # MEDIDAS: Unused e Used by unused
    status_remover_medidas = ['Unused', 'Used by unused', 'unused', 'used by unused']
    medidas_sujeira = df[
        (df[coluna_tipo].str.contains('Measure', case=False, na=False)) & 
        (df[coluna_status].isin(status_remover_medidas))
    ]

    # COLUNAS: Apenas Unused para não quebrar relacionamentos/SortBy
    colunas_sujeira = df[
        (df[coluna_tipo].str.contains('Column', case=False, na=False)) & 
        (df[coluna_status].str.lower() == 'unused')
    ]

    sujeira_df = pd.concat([medidas_sujeira, colunas_sujeira])
    
    # Criar mapa de alvos por tabela
    mapa_sujeira = {}
    for _, row in sujeira_df.iterrows():
        tabela = str(row[coluna_tabela]).strip()
        if tabela not in mapa_sujeira:
            mapa_sujeira[tabela] = []
        mapa_sujeira[tabela].append(str(row[coluna_nome]).strip())

    log_remocao = []

    # 3. Processar Arquivos TMDL
    if not os.path.exists(PBIP_TABLES_PATH):
        print(f"Pasta TMDL não encontrada: {PBIP_TABLES_PATH}")
        return

    for filename in os.listdir(PBIP_TABLES_PATH):
        if not filename.endswith('.tmdl'): continue
        
        table_name = filename.replace('.tmdl', '')
        
        # PROTEÇÃO: Ignora tabelas protegidas e tabelas automáticas de data
        if table_name in TABELAS_FIXAS_EXCECAO or \
           table_name.startswith(('LocalDateTable_', 'DateTableTemplate_')):
            continue

        if table_name in mapa_sujeira:
            file_path = os.path.join(PBIP_TABLES_PATH, filename)
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            novas_linhas = []
            skip_block = False
            alvos = mapa_sujeira[table_name]

            for line in lines:
                clean_line = line.strip()
                
                # Identifica início de novo objeto
                if clean_line.startswith(('column ', 'measure ', 'partition ', 'table ', 'hierarchy ')):
                    # Extrai o nome limpando aspas e o sinal de '='
                    parts = clean_line.split(' ', 1)
                    nome_cru = parts[1].split('=')[0].strip()
                    current_name = nome_cru.replace("'", "")
                    
                    if current_name in alvos:
                        skip_block = True
                        log_remocao.append(f"Removido: [{table_name}] -> {current_name}")
                    else:
                        skip_block = False

                # Se a linha encosta na margem (0 tabs) e não é vazia, interrompe o skip
                num_tabs = len(line) - len(line.lstrip('\t'))
                if skip_block and num_tabs == 0 and clean_line != "":
                    skip_block = False

                if not skip_block:
                    novas_linhas.append(line)

            # Salva o arquivo limpo
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(novas_linhas)
            print(f"✅ Tabela [{table_name}] limpa.")

    # 4. Gravar Log
    with open('log_limpeza.txt', 'w', encoding='utf-8') as f_log:
        f_log.write(f"Total de itens removidos: {len(log_remocao)}\n" + "="*30 + "\n")
        f_log.write("\n".join(log_remocao))
    
    print(f"\n -- Limpeza Finalizada! {len(log_remocao)} itens removidos.")
    print(f"-- Verifique o arquivo 'log_limpeza.txt' para detalhes.")

if __name__ == "__main__":
    clean_tmdl()