import pandas as pd
import os
from datetime import datetime

pasta = r"C:\Users\x13376866\Desktop\frota_sgtp_sgtm_simetro"
arquivo_excel = os.path.join(pasta, "relatorio_comparativo_frotas.xlsx")
arquivo_txt = os.path.join(pasta, "relatorio_procedimento_frotas.txt")

# 1. Carregamento dos Dados
df_sgtm = pd.read_csv(os.path.join(pasta, "frota_sgtm"), low_memory=False)
df_sgtp = pd.read_csv(os.path.join(pasta, "frota_sgtp.csv"), sep=';', encoding='latin1', low_memory=False)

try:
    df_simetro = pd.read_csv(
        os.path.join(pasta, "frota_simetro.csv"), 
        skiprows=1, 
        sep=';', 
        encoding='latin1', 
        on_bad_lines='skip'
    )
except Exception:
    df_simetro = pd.read_csv(
        os.path.join(pasta, "frota_simetro.csv"), 
        skiprows=1, 
        sep=',', 
        encoding='latin1', 
        on_bad_lines='skip'
    )

# Identificação da coluna de Placa em cada base
def get_col(df, termo):
    cols = [c for c in df.columns if termo.lower() in str(c).lower()]
    return cols[0] if cols else None

col_p_sgtm = get_col(df_sgtm, 'placa')
col_p_sgtp = get_col(df_sgtp, 'placa')
col_p_simetro = get_col(df_simetro, 'placa')

# Tratamento e Limpeza das Placas
def limpar_campo(serie):
    if serie is None:
        return pd.Series(dtype=str)
    return (
        serie.dropna()
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace('-', '', regex=False)
        .str.replace(' ', '', regex=False)
        .str.replace('.', '', regex=False)
    )

df_sgtm['PLACA_LIMPA'] = limpar_campo(df_sgtm[col_p_sgtm]) if col_p_sgtm else ''
df_sgtp['PLACA_LIMPA'] = limpar_campo(df_sgtp[col_p_sgtp]) if col_p_sgtp else ''
df_simetro['PLACA_LIMPA'] = limpar_campo(df_simetro[col_p_simetro]) if col_p_simetro else ''

# Deduplicação por Placa
df_sgtm_u = df_sgtm.drop_duplicates(subset=['PLACA_LIMPA']).copy()
df_sgtp_u = df_sgtp.drop_duplicates(subset=['PLACA_LIMPA']).copy()
df_simetro_u = df_simetro.drop_duplicates(subset=['PLACA_LIMPA']).copy()

p_sgtm = set(df_sgtm_u['PLACA_LIMPA'][df_sgtm_u['PLACA_LIMPA'] != ''])
p_sgtp = set(df_sgtp_u['PLACA_LIMPA'][df_sgtp_u['PLACA_LIMPA'] != ''])
p_simetro = set(df_simetro_u['PLACA_LIMPA'][df_simetro_u['PLACA_LIMPA'] != ''])

# Operações de Conjunto (Cruzamentos)
em_todos = p_sgtm & p_sgtp & p_simetro
apenas_sgtm = p_sgtm - p_sgtp - p_simetro
apenas_sgtp = p_sgtp - p_sgtm - p_simetro
apenas_simetro = p_simetro - p_sgtm - p_sgtp

sgtm_sgtp_sem_simetro = (p_sgtm & p_sgtp) - p_simetro
sgtm_simetro_sem_sgtp = (p_sgtm & p_simetro) - p_sgtp
sgtp_simetro_sem_sgtm = (p_sgtp & p_simetro) - p_sgtm

# -----------------------------------------------------------------------------
# 2. ESTUDO DETALHADO DOS VEÍCULOS EXCLUSIVOS DO SGTM
# -----------------------------------------------------------------------------
df_exc_sgtm = df_sgtm_u[df_sgtm_u['PLACA_LIMPA'].isin(apenas_sgtm)].copy()

# Identificação de campos de status e data de saída no SGTM
col_saida = get_col(df_exc_sgtm, 'data_saida') or get_col(df_exc_sgtm, 'saida')
col_ano = get_col(df_exc_sgtm, 'ano_carroceria') or get_col(df_exc_sgtm, 'ano_chassi') or get_col(df_exc_sgtm, 'ano')

# Classificação entre Baixados (Com data de saída) e Ativos/Sem Saída Cadastrada
if col_saida and col_saida in df_exc_sgtm.columns:
    # Trata datas inválidas/padrão ex: 1899-12-30 ou nulas
    saida_limpa = df_exc_sgtm[col_saida].astype(str).str.strip()
    tem_saida = saida_limpa.notna() & (~saida_limpa.isin(['nan', '', 'None', 'NULL'])) & (~saida_limpa.str.startswith('1899'))
    
    df_exc_sgtm['STATUS_SGTM'] = 'Ativo / Sem Data de Saída'
    df_exc_sgtm.loc[tem_saida, 'STATUS_SGTM'] = 'Baixado / Com Data de Saída'
    
    total_baixados = int(tem_saida.sum())
    total_sem_saida = int((~tem_saida).sum())
else:
    df_exc_sgtm['STATUS_SGTM'] = 'Não Identificado'
    total_baixados = 0
    total_sem_saida = len(df_exc_sgtm)

# Perfil do Ano de Fabricação/Carroceria
if col_ano and col_ano in df_exc_sgtm.columns:
    anos_validos = pd.to_numeric(df_exc_sgtm[col_ano], errors='coerce').dropna()
    ano_min = int(anos_validos.min()) if not anos_validos.empty else 'N/A'
    ano_max = int(anos_validos.max()) if not anos_validos.empty else 'N/A'
    ano_medio = round(float(anos_validos.mean()), 1) if not anos_validos.empty else 'N/A'
else:
    ano_min, ano_max, ano_medio = 'N/A', 'N/A', 'N/A'

# -----------------------------------------------------------------------------
# 3. TABELA RESUMO PARA O EXCEL
# -----------------------------------------------------------------------------
df_resumo = pd.DataFrame([
    {"Categoria / Status": "Presentes nos 3 Sistemas", "Quantidade de Placas": len(em_todos), "Descrição": "Frota comum operando convergentemente nos 3 sistemas"},
    {"Categoria / Status": "Exclusivas do SGTM (Total)", "Quantidade de Placas": len(apenas_sgtm), "Descrição": "Veículos presentes apenas no SGTM"},
    {"Categoria / Status": "  -> SGTM: Com Data de Saída (Baixados/Histórico)", "Quantidade de Placas": total_baixados, "Descrição": "Veículos inativos com registro formal de baixa no SGTM"},
    {"Categoria / Status": "  -> SGTM: Sem Data de Saída (Pendentes/Inconsistência)", "Quantidade de Placas": total_sem_saida, "Descrição": "Veículos no SGTM sem data de saída, mas ausentes do SGTP/SIMETRO"},
    {"Categoria / Status": "Exclusivas do SGTP", "Quantidade de Placas": len(apenas_sgtp), "Descrição": "Veículos cadastrados apenas no SGTP"},
    {"Categoria / Status": "Exclusivas do SIMETRO", "Quantidade de Placas": len(apenas_simetro), "Descrição": "Veículos cadastrados apenas no SIMETRO"},
    {"Categoria / Status": "Presente em SGTM e SGTP (sem SIMETRO)", "Quantidade de Placas": len(sgtm_sgtp_sem_simetro), "Descrição": "Presentes no SGTM e SGTP, ausentes no SIMETRO"},
    {"Categoria / Status": "Presente em SGTM e SIMETRO (sem SGTP)", "Quantidade de Placas": len(sgtm_simetro_sem_sgtp), "Descrição": "Presentes no SGTM e SIMETRO, ausentes no SGTP"},
    {"Categoria / Status": "Presente em SGTP e SIMETRO (sem SGTM)", "Quantidade de Placas": len(sgtp_simetro_sem_sgtm), "Descrição": "Presentes no SGTP e SIMETRO, ausentes no SGTM"},
    {"Categoria / Status": "Total Placas Únicas - SGTM", "Quantidade de Placas": len(p_sgtm), "Descrição": "Total geral de placas no SGTM"},
    {"Categoria / Status": "Total Placas Únicas - SGTP", "Quantidade de Placas": len(p_sgtp), "Descrição": "Total geral de placas no SGTP"},
    {"Categoria / Status": "Total Placas Únicas - SIMETRO", "Quantidade de Placas": len(p_simetro), "Descrição": "Total geral de placas no SIMETRO"}
])

# -----------------------------------------------------------------------------
# 4. EXPORTAÇÃO EXCEL
# -----------------------------------------------------------------------------
with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
    df_resumo.to_excel(writer, sheet_name='Resumo_Geral', index=False)
    df_sgtm_u[df_sgtm_u['PLACA_LIMPA'].isin(em_todos)].to_excel(writer, sheet_name='Presente_nos_3', index=False)
    df_exc_sgtm.to_excel(writer, sheet_name='Estudo_Exclusivas_SGTM', index=False)
    df_sgtp_u[df_sgtp_u['PLACA_LIMPA'].isin(apenas_sgtp)].to_excel(writer, sheet_name='Apenas_SGTP', index=False)
    df_simetro_u[df_simetro_u['PLACA_LIMPA'].isin(apenas_simetro)].to_excel(writer, sheet_name='Apenas_SIMETRO', index=False)

# -----------------------------------------------------------------------------
# 5. GERADOR DO RELATÓRIO EM TEXTO (PROCEDIMENTOS + ESTUDO SGTM)
# -----------------------------------------------------------------------------
relatorio_texto = f"""================================================================================
RELATÓRIO PROCEDIMENTAL E AUDITORIA COMPARATIVA DE FROTAS
SISTEMAS ANALISADOS: SGTM | SGTP | SIMETRO
Data de Processamento: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
================================================================================

1. DIAGNÓSTICO E ESTRUTURA DOS ARQUIVOS FONTE
--------------------------------------------------------------------------------
- Arquivo 'frota_sgtm':
    * Formato: CSV (sem extensão) | Delimitador: Vírgula (,) | Encoding: UTF-8
    * Total de Linhas: {len(df_sgtm)} | Colunas: {len(df_sgtm.columns)}
    * Coluna de Placa Utilizada: '{col_p_sgtm}'

- Arquivo 'frota_sgtp.csv':
    * Formato: CSV | Delimitador: Ponto e vírgula (;) | Encoding: Latin1
    * Total de Linhas: {len(df_sgtp)} | Colunas: {len(df_sgtp.columns)}
    * Coluna de Placa Utilizada: '{col_p_sgtp}'

- Arquivo 'frota_simetro.csv':
    * Formato: CSV | Delimitador: Ponto e vírgula (;) | Encoding: Latin1
    * Tratamento Especial: Pulo da 1ª linha de metragem/cabeçalho (skiprows=1)
    * Total de Linhas: {len(df_simetro)} | Colunas: {len(df_simetro.columns)}
    * Coluna de Placa Utilizada: '{col_p_simetro}'

2. PROCEDIMENTOS DE TRATAMENTO E HIGIENIZAÇÃO DE DADOS
--------------------------------------------------------------------------------
a) Remoção de hífens, pontos e espaços das placas para permitir o cruzamento.
b) Padronização de todas as placas em letras maiúsculas.
c) Deduplicação de registros dentro de cada base individual pela Placa.

3. RESUMO GERAL DAS FROTAS E CRUZAMENTOS
--------------------------------------------------------------------------------
- Total de Placas Únicas no SGTM:    {len(p_sgtm):>6}
- Total de Placas Únicas no SGTP:    {len(p_sgtp):>6}
- Total de Placas Únicas no SIMETRO: {len(p_simetro):>6}

- Veículos Presentes nos 3 Sistemas:                {len(em_todos):>6}
- Veículos Exclusivos do SGTM (Apenas SGTM):        {len(apenas_sgtm):>6}
- Veículos Exclusivos do SGTP (Apenas SGTP):        {len(apenas_sgtp):>6}
- Veículos Exclusivos do SIMETRO (Apenas SIMETRO):  {len(apenas_simetro):>6}

- Presentes em SGTM e SGTP (Ausentes no SIMETRO):    {len(sgtm_sgtp_sem_simetro):>6}
- Presentes em SGTM e SIMETRO (Ausentes no SGTP):    {len(sgtm_simetro_sem_sgtp):>6}
- Presentes em SGTP e SIMETRO (Ausentes no SGTM):    {len(sgtp_simetro_sem_sgtm):>6}

4. ESTUDO DETALHADO DA FROTA EXCLUSIVA DO SGTM ({len(apenas_sgtm)} PLACAS)
--------------------------------------------------------------------------------
Objetivo: Compreender a natureza dos 1.320 veículos que existem apenas no SGTM.

a) Status Operacional / Registro de Saída:
   * Veículos Baixados (Com Data de Saída Cadastrada): {total_baixados:>5} ({round(total_baixados/len(apenas_sgtm)*100, 1)}%)
     -> Correspondem a veículos antigos desativados/substituídos ao longo dos anos
        que permanecem gravados no histórico do SGTM.
   * Veículos sem Registro de Saída (Possível Inconsistência): {total_sem_saida:>5} ({round(total_sem_saida/len(apenas_sgtm)*100, 1)}%)
     -> Veículos sem data formal de saída no SGTM, porém ausentes nos cadastros
        operacionais do SGTP e SIMETRO (requerem auditoria de baixa).

b) Perfil Temporal (Ano de Carroceria / Chassi):
   * Ano Mais Antigo Encontrado: {ano_min}
   * Ano Mais Recente Encontrado: {ano_max}
   * Ano Médio da Frota Exclusiva SGTM: {ano_medio}

5. ARQUIVOS GERADOS NA PASTA
--------------------------------------------------------------------------------
- Planilha Excel: {os.path.basename(arquivo_excel)}
    * Abas: Resumo_Geral, Presente_nos_3, Estudo_Exclusivas_SGTM, 
            Apenas_SGTP, Apenas_SIMETRO.
- Relatório de Texto: {os.path.basename(arquivo_txt)}
================================================================================
"""

with open(arquivo_txt, "w", encoding="utf-8") as f:
    f.write(relatorio_texto)

print("\n=======================================================")
print("Processamento e Estudo do SGTM concluídos com sucesso!")
print(f"Excel gerado: {arquivo_excel}")
print(f"Relatório TXT gerado: {arquivo_txt}")
print("=======================================================\n")