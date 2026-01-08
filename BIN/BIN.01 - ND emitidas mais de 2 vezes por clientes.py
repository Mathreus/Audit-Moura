# AUDITORES, NA LINHA 34 DEVEMOS ALTERAR O CÓDIGO DO DISTRIBUIDOR QUE IRÁ SER AUDITADO!!!
# NA LINHA 35 DEVEMOS ALTERAR AS DATAS PARA O ESCOPO A SER AUDITADO!!!
# O MESMO DEVE SER FEITO NAS LINHAS 89 E 90
# NA LINHA 141 DEVEMOS INSERIR O CAMINHO ONDE A PLANILHA GERADA SERÁ SALVA!!!

import pyodbc
import pandas as pd
import os
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Defina as informações de conexão
server = 'DCMDWF01A.MOURA.INT'
database = 'ax'
username = 'uAuditoria'
password = '@ud!t0$!@202&22'
driver = '{SQL Server}'  # Corrigido: precisa estar entre chaves

# Construa a string de conexão
connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Queries SQL
QUERY_SINTETICO = """
WITH ClientesMultiplosTND AS (
    SELECT 
        COD_CLIENTE,
        NOME_CLIENTE,
        COUNT(DISTINCT NOTA_FISCAL) as qtd_notas_fiscais,
        SUM(VALOR_TITULO) AS VALOR_TOTAL
    FROM 
        VW_AUDIT_RM_TRANSACOES_FECHADAS
    WHERE 
        COD_ESTABELECIMENTO IN ('R351', 'R352') 
        AND DATA_TRANSACAO BETWEEN '2025-07-01' AND '2025-12-31'
        AND PERFIL_LANCAMENTO = 'TND'
    GROUP BY 
        COD_CLIENTE, NOME_CLIENTE
    HAVING 
        COUNT(DISTINCT NOTA_FISCAL) > 2
)
SELECT
    'Dinil' AS COD_ESTABELECIMENTO,
    c.COD_CLIENTE,    
    c.NOME_CLIENTE,
    c.qtd_notas_fiscais AS TOTAL_NF_DIFERENTES,
    c.VALOR_TOTAL
FROM
    ClientesMultiplosTND c
ORDER BY
    c.qtd_notas_fiscais DESC,
    c.VALOR_TOTAL DESC
"""

QUERY_ANALITICO = """
WITH ClientesMultiplosTND AS (
    SELECT 
        COD_CLIENTE,
        COUNT(DISTINCT NOTA_FISCAL) as qtd_notas_fiscais,
        SUM(VALOR_TITULO) AS VALOR_TOTAL
    FROM 
        VW_AUDIT_RM_TRANSACOES_FECHADAS
    WHERE 
        COD_ESTABELECIMENTO IN ('R351', 'R352')
        AND DATA_TRANSACAO BETWEEN '2025-07-01' AND '2025-12-31'
        AND PERFIL_LANCAMENTO = 'TND'
    GROUP BY 
        COD_CLIENTE
    HAVING 
        COUNT(DISTINCT NOTA_FISCAL) > 2
)
SELECT
    v.COD_ESTABELECIMENTO,
    v.COD_CLIENTE,    
    v.NOME_CLIENTE,
    v.DATA_TRANSACAO,
    v.DATA_VENCIMENTO,
    v.PERFIL_LANCAMENTO,
    v.NOTA_FISCAL,
    v.COMPROVANTE,
    v.PARCELA,
    v.VALOR_PARCELA,
    v.VALOR_TITULO
FROM
    VW_AUDIT_RM_TRANSACOES_FECHADAS v
INNER JOIN 
    ClientesMultiplosTND c ON v.COD_CLIENTE = c.COD_CLIENTE
WHERE
    v.COD_ESTABELECIMENTO IN ('R351', 'R352')
    AND v.DATA_TRANSACAO BETWEEN '2025-07-01' AND '2025-12-31'
    AND v.PERFIL_LANCAMENTO = 'TND'
ORDER BY
    v.COD_CLIENTE,
    v.NOTA_FISCAL,
    v.DATA_TRANSACAO
"""

try:
    # Conecte-se ao banco de dados
    print("Conectando ao banco de dados...")
    conexao = pyodbc.connect(connection_string)
    print("✅ Conexão estabelecida com sucesso!")
    
    # Executar as queries usando pandas
    print("\n🔍 Executando query sintética...")
    df_sintetico = pd.read_sql(QUERY_SINTETICO, conexao)
    print(f"✅ Query sintética executada: {len(df_sintetico)} clientes encontrados")
    
    print("\n🔍 Executando query analítica...")
    df_analitico = pd.read_sql(QUERY_ANALITICO, conexao)
    print(f"✅ Query analítica executada: {len(df_analitico)} transações encontradas")
    
    # Fechar a conexão
    conexao.close()
    print("\n🔒 Conexão com o banco de dados encerrada.")
    
    # Formatar os DataFrames - APENAS DATAS (mantendo valores numéricos)
    def formatar_dataframe(df, tipo):
        """Aplica formatação específica aos DataFrames - APENAS DATAS"""
        if tipo == 'analitico':
            # Formatar APENAS colunas de data (mantendo valores numéricos)
            colunas_data = ['DATA_TRANSACAO', 'DATA_VENCIMENTO']
            for col in colunas_data:
                if col in df.columns and not df[col].empty:
                    try:
                        df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y')
                    except:
                        pass
        return df
    
    # Aplicar formatação APENAS de datas
    df_sintetico = formatar_dataframe(df_sintetico, 'sintetico')
    df_analitico = formatar_dataframe(df_analitico, 'analitico')
    
    # Exportar para Excel
    # Gerar nome do arquivo com timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"Clientes_TND_R241_{timestamp}.xlsx"
    
    # Definir o caminho específico
    caminho_base = r"C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2026\03. Automações\Validações"  # INSIRA AQUI O CAMINHO QUE DESEJA SALVAR O RESULTADO!!!
    caminho_completo = os.path.join(caminho_base, nome_arquivo)
    
    # Verificar se o diretório existe, se não, criar
    if not os.path.exists(caminho_base):
        os.makedirs(caminho_base)
        print(f"\n📁 Diretório criado: {caminho_base}")
    
    # Exportar para Excel com duas abas
    print("\n📊 Gerando arquivo Excel...")
    try:
        with pd.ExcelWriter(caminho_completo, engine='openpyxl') as writer:
            # Escrever aba sintética
            df_sintetico.to_excel(writer, sheet_name='Resumo_Sintético', index=False)
            
            # Escrever aba analítica
            df_analitico.to_excel(writer, sheet_name='Detalhes_Analítico', index=False)
            
            # Ajustar largura das colunas E aplicar formatação numérica
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                
                # Identificar colunas numéricas para formatação
                colunas_valor = []
                if sheet_name == 'Resumo_Sintético':
                    colunas_valor = ['VALOR_TOTAL']
                elif sheet_name == 'Detalhes_Analítico':
                    colunas_valor = ['VALOR_PARCELA', 'VALOR_TITULO']
                
                # Mapear nomes de colunas para letras de coluna
                col_letters = {}
                for idx, col in enumerate(df_sintetico.columns if sheet_name == 'Resumo_Sintético' else df_analitico.columns, start=1):
                    col_letters[col] = chr(64 + idx) if idx <= 26 else chr(64 + (idx // 26)) + chr(64 + (idx % 26))
                
                # Aplicar formatação numérica às colunas de valor
                for col_name in colunas_valor:
                    if col_name in col_letters:
                        col_letter = col_letters[col_name]
                        # Formatar como moeda brasileira
                        for cell in worksheet[col_letter]:
                            if cell.row > 1:  # Ignorar cabeçalho
                                cell.number_format = 'R$ #,##0.00'
                
                # Ajustar largura das colunas
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Arquivo Excel gerado com sucesso: {caminho_completo}")
        print(f"\n📊 RESUMO DO RELATÓRIO:")
        print(f"   - Estabelecimento: R351 e R352")
        print(f"   - Período: 01/07/2025 a 31/12/2025")
        print(f"   - Perfil Lançamento: TND")
        print(f"   - Critério: Clientes com >2 notas fiscais diferentes")
        print(f"   - Total de clientes: {len(df_sintetico)}")
        print(f"   - Total de transações: {len(df_analitico)}")
        print(f"   - Campos numéricos mantidos em formato numérico")
        
    except Exception as e:
        print(f"❌ Erro ao exportar para Excel: {e}")

except pyodbc.Error as e:
    print(f"❌ Erro na conexão com o banco de dados: {e}")
except Exception as e:
    print(f"❌ Erro no processamento: {e}")
