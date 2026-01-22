import pyodbc
import pandas as pd
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')

class CalculoIndiceDevolucao:
    def __init__(self):
        """
        Inicializa com as credenciais fornecidas
        """
        self.server = 'DCMDWF01A.MOURA.INT'
        self.database = 'ax'
        self.username = 'uAuditoria'
        self.password = '@ud!t0$!@202&22'
        self.driver = '{SQL Server}'  # Corrigido: usando chaves
        self.conn = None
        
        # Definir o caminho específico para salvar
        self.caminho_base = r'C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2026\03. Automações\Validações'
        
    def conectar_banco(self):
        """Estabelece conexão com o banco de dados"""
        try:
            connection_string = (
                f'DRIVER={self.driver};'
                f'SERVER={self.server};'
                f'DATABASE={self.database};'
                f'UID={self.username};'
                f'PWD={self.password};'
            )
            
            self.conn = pyodbc.connect(connection_string)
            print("Conexão estabelecida com sucesso!")
            return True
        except Exception as e:
            print(f"Erro ao conectar ao banco de dados: {e}")
            return False
    
    def executar_query(self, query):
        """Executa uma query SQL e retorna um DataFrame"""
        try:
            df = pd.read_sql_query(query, self.conn)
            return df
        except Exception as e:
            print(f"Erro ao executar query: {e}")
            return None
    
    def obter_dados_devolucao(self, data_inicio='2025-07-01', data_fim='2025-12-31'):
        """Obtém os dados de devolução por vendedor"""
        query_devolucao = f"""
        SELECT
            COD_ESTABELECIMENTO,
            VENDEDOR,
            SUM(QUANTIDADE) AS TOTAL_QTD_DEVOLVIDA,
            SUM(VALOR) AS VALOR_DEVOLUCAO,
            COUNT(*) AS QTD_NOTAS_DEVOLVIDAS
        FROM 
            VW_AUDIT_RM_ORDENS_VENDA
        WHERE
            COD_ESTABELECIMENTO IN ('R351', 'R352')
            AND DATA_NOTA_FISCAL BETWEEN '{data_inicio}' AND '{data_fim}' 
            AND PARA_FATURAMENTO = 'SIM'
            AND NUM_NOTA_FISCAL NOT LIKE '%EST%'
            AND CFOP IN ('1.201', '1.202', '1.203', '1.204', '1.410', '1.411', '1.553', '1.660', '1.661', '1.662', 
                '2.201', '2.202', '2.203', '2.204', '2.410', '2.411', '2.553', '2.660', '2.661', '2.662',
                '3.201', '3.202', '3.211', '3.553')
        GROUP BY
            COD_ESTABELECIMENTO,
            VENDEDOR
        HAVING 
            COUNT(*) <> 0
        ORDER BY
            COD_ESTABELECIMENTO,
            VENDEDOR,
            VALOR_DEVOLUCAO ASC
        """
        
        print("Executando query de devolução...")
        df_devolucao = self.executar_query(query_devolucao)
        return df_devolucao
    
    def obter_dados_faturamento(self, data_inicio='2025-07-01', data_fim='2025-12-31'):
        """Obtém os dados de faturamento por vendedor"""
        query_faturamento = f"""
        SELECT
            COD_ESTABELECIMENTO,
            VENDEDOR,
            SUM(QUANTIDADE) AS TOTAL_QTD_FATURADA,
            ABS(SUM(VALOR)) AS VALOR_FATURAMENTO,
            COUNT(*) AS QTD_NOTAS_FATURADAS
        FROM 
            VW_AUDIT_RM_ORDENS_VENDA
        WHERE 
            COD_ESTABELECIMENTO IN ('R351', 'R352')  
            AND DATA_NOTA_FISCAL BETWEEN '{data_inicio}' AND '{data_fim}'  
            AND PARA_FATURAMENTO = 'Sim'
            AND NUM_NOTA_FISCAL NOT LIKE '%EST%'
            AND CFOP IN ('5.100', '5.101', '5.102', '5.103', '5.104', '5.105', '5.106', '5.109', '5.110', '5.111', 
                        '5.112', '5.113', '5.114', '5.115', '5.116', '5.117', '5.118', '5.119', '5.120', '5.122', 
                        '5.123', '5.250','5.251', '5.252', '5.253', '5.254', '5.255', '5.256', '5.257', '5.258', 
                        '5.401', '5.402', '5.403', '5.405', '5.651', '5.652', '5.653', '5.654', '5.655', '5.656',
                        '5.667', '6.101', '6.102', '6.103','6.104', '6.105', '6.106', '6.107', '6.108', '6.109',
                        '6.110', '6.111', '6.112', '6.113', '6.114', '6.115', '6.116', '6.117', '6.118', '6.119',
                        '6.120', '6.122', '6.123', '6.250', '6.251', '6.252', '6.253', '6.254', '6.255', '6.256',
                        '6.257', '6.258', '6.401', '6.402', '6.403', '6.404', '6.651', '6.652', '6.653', '6.654',
                        '6.655', '6.656', '6.667', '7.100', '7.101', '7.102', '7.105', '7.106','7.127', '7.250', 
                        '7.251', '7.651', '7.654', '7.667')
        GROUP BY    
            COD_ESTABELECIMENTO,
            VENDEDOR
        ORDER BY    
            COD_ESTABELECIMENTO,
            VENDEDOR,
            VALOR_FATURAMENTO ASC
        """
        
        print("Executando query de faturamento...")
        df_faturamento = self.executar_query(query_faturamento)
        return df_faturamento
    
    def calcular_indice_devolucao(self, df_devolucao, df_faturamento):
        """Calcula o índice de devolução: VALOR_DEVOLVIDO / VALOR_FATURADO"""
        if df_devolucao is None or df_faturamento is None:
            print("Erro: Dados não disponíveis para cálculo")
            return None
        
        # Criar chave única para merge (agora só com estabelecimento e vendedor)
        df_devolucao['chave'] = df_devolucao['COD_ESTABELECIMENTO'] + '_' + df_devolucao['VENDEDOR'].astype(str)
        df_faturamento['chave'] = df_faturamento['COD_ESTABELECIMENTO'] + '_' + df_faturamento['VENDEDOR'].astype(str)
        
        # Realizar merge dos dataframes
        df_consolidado = pd.merge(
            df_faturamento,
            df_devolucao[['chave', 'VALOR_DEVOLUCAO', 'TOTAL_QTD_DEVOLVIDA', 'QTD_NOTAS_DEVOLVIDAS']],
            on='chave',
            how='left'
        )
        
        # Preencher valores NaN com 0
        df_consolidado['VALOR_DEVOLUCAO'] = df_consolidado['VALOR_DEVOLUCAO'].fillna(0)
        df_consolidado['TOTAL_QTD_DEVOLVIDA'] = df_consolidado['TOTAL_QTD_DEVOLVIDA'].fillna(0)
        df_consolidado['QTD_NOTAS_DEVOLVIDAS'] = df_consolidado['QTD_NOTAS_DEVOLVIDAS'].fillna(0)
        
        # Calcular índice de devolução
        df_consolidado['INDICE_DEVOLUCAO'] = df_consolidado.apply(
            lambda row: row['VALOR_DEVOLUCAO'] / row['VALOR_FATURAMENTO'] 
            if row['VALOR_FATURAMENTO'] > 0 else 0,
            axis=1
        )
        
        # Calcular percentual de devolução
        df_consolidado['PERCENTUAL_DEVOLUCAO'] = df_consolidado['INDICE_DEVOLUCAO'] * 100
        
        # Calcular outros indicadores
        df_consolidado['PERCENTUAL_QTD_DEVOLVIDA'] = df_consolidado.apply(
            lambda row: (row['TOTAL_QTD_DEVOLVIDA'] / row['TOTAL_QTD_FATURADA'] * 100)
            if row['TOTAL_QTD_FATURADA'] > 0 else 0,
            axis=1
        )
        
        # Ordenar por maior índice de devolução
        df_consolidado = df_consolidado.sort_values('INDICE_DEVOLUCAO', ascending=False)
        
        # Remover coluna auxiliar
        df_consolidado = df_consolidado.drop('chave', axis=1)
        
        # Adicionar coluna de ranking
        df_consolidado['RANKING'] = range(1, len(df_consolidado) + 1)
        
        # Reorganizar colunas
        colunas = ['RANKING', 'COD_ESTABELECIMENTO', 'VENDEDOR', 
                  'VALOR_FATURAMENTO', 'QTD_NOTAS_FATURADAS', 'TOTAL_QTD_FATURADA',
                  'VALOR_DEVOLUCAO', 'QTD_NOTAS_DEVOLVIDAS', 'TOTAL_QTD_DEVOLVIDA',
                  'INDICE_DEVOLUCAO', 'PERCENTUAL_DEVOLUCAO', 'PERCENTUAL_QTD_DEVOLVIDA']
        
        # Manter apenas as colunas que existem
        colunas = [col for col in colunas if col in df_consolidado.columns]
        df_consolidado = df_consolidado[colunas]
        
        return df_consolidado
    
    def exportar_resultados(self, df_resultado, formato='excel'):
        """Exporta os resultados para Excel ou CSV no caminho especificado"""
        try:
            # Verificar se o diretório existe, se não, criar
            if not os.path.exists(self.caminho_base):
                print(f"Criando diretório: {self.caminho_base}")
                os.makedirs(self.caminho_base)
            
            # Criar nome do arquivo com data e hora
            data_atual = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            if formato.lower() == 'excel':
                nome_arquivo = f'indice_devolucao_vendedores_{data_atual}.xlsx'
                caminho_completo = os.path.join(self.caminho_base, nome_arquivo)
                
                # Garantir que os nomes das colunas sejam strings válidas para Excel
                df_resultado.columns = [str(col).replace('/', '_').replace('[', '').replace(']', '') 
                                       for col in df_resultado.columns]
                
                # Exportar para Excel
                with pd.ExcelWriter(caminho_completo, engine='openpyxl') as writer:
                    # Planilha principal com todos os dados
                    df_resultado.to_excel(writer, sheet_name='Indice_Devolucao', index=False)
                    
                    # Criar uma planilha com resumo
                    resumo = self._criar_resumo(df_resultado)
                    resumo.to_excel(writer, sheet_name='Resumo', index=False)
                    
                    # Planilha com top 20 maiores índices
                    top_20 = df_resultado.head(20).copy()
                    top_20.to_excel(writer, sheet_name='Top_20_Maiores', index=False)
                    
                    # Planilha com análise por estabelecimento
                    analise_estabelecimento = self._criar_analise_estabelecimento(df_resultado)
                    analise_estabelecimento.to_excel(writer, sheet_name='Analise_Estabelecimento', index=False)
                
                print(f"Resultados exportados para: {caminho_completo}")
                
                # Também criar um arquivo CSV para fácil visualização
                nome_csv = f'indice_devolucao_vendedores_{data_atual}.csv'
                caminho_csv = os.path.join(self.caminho_base, nome_csv)
                df_resultado.to_csv(caminho_csv, index=False, sep=';', decimal=',', encoding='utf-8')
                print(f"Arquivo CSV também criado em: {caminho_csv}")
                
            else:
                nome_arquivo = f'indice_devolucao_vendedores_{data_atual}.csv'
                caminho_completo = os.path.join(self.caminho_base, nome_arquivo)
                df_resultado.to_csv(caminho_completo, index=False, sep=';', decimal=',', encoding='utf-8')
                print(f"Resultados exportados para: {caminho_completo}")
            
            return caminho_completo
            
        except Exception as e:
            print(f"Erro ao exportar resultados: {e}")
            # Tentar salvar no diretório atual como fallback
            try:
                data_atual = datetime.now().strftime('%Y%m%d_%H%M%S')
                nome_arquivo = f'indice_devolucao_vendedores_{data_atual}.xlsx'
                df_resultado.to_excel(nome_arquivo, index=False)
                print(f"Resultados exportados para o diretório atual: {nome_arquivo}")
                return nome_arquivo
            except:
                print("Não foi possível exportar os resultados.")
                return None
    
    def _criar_resumo(self, df_resultado):
        """Cria um DataFrame com estatísticas resumidas"""
        resumo_data = {
            'Metrica': [
                'Total de Vendedores',
                'Valor Total Faturado (R$)',
                'Valor Total Devolvido (R$)',
                'Percentual Total de Devolução (%)',
                'Média do Índice de Devolução',
                'Mediana do Índice de Devolução',
                'Máximo Índice de Devolução',
                'Mínimo Índice de Devolução',
                'Média Percentual Devolução (%)',
                'Vendedores com Devolução > 0',
                'Vendedores com Devolução = 0',
                'Quantidade Total Faturada',
                'Quantidade Total Devolvida',
                'Percentual Qtd Devolvida (%)',
                'Notas Faturadas',
                'Notas Devolvidas'
            ],
            'Valor': [
                df_resultado['VENDEDOR'].nunique(),
                df_resultado['VALOR_FATURAMENTO'].sum(),
                df_resultado['VALOR_DEVOLUCAO'].sum(),
                (df_resultado['VALOR_DEVOLUCAO'].sum() / df_resultado['VALOR_FATURAMENTO'].sum() * 100) if df_resultado['VALOR_FATURAMENTO'].sum() > 0 else 0,
                df_resultado['INDICE_DEVOLUCAO'].mean(),
                df_resultado['INDICE_DEVOLUCAO'].median(),
                df_resultado['INDICE_DEVOLUCAO'].max(),
                df_resultado['INDICE_DEVOLUCAO'].min(),
                df_resultado['PERCENTUAL_DEVOLUCAO'].mean(),
                (df_resultado['VALOR_DEVOLUCAO'] > 0).sum(),
                (df_resultado['VALOR_DEVOLUCAO'] == 0).sum(),
                df_resultado['TOTAL_QTD_FATURADA'].sum(),
                df_resultado['TOTAL_QTD_DEVOLVIDA'].sum(),
                (df_resultado['TOTAL_QTD_DEVOLVIDA'].sum() / df_resultado['TOTAL_QTD_FATURADA'].sum() * 100) if df_resultado['TOTAL_QTD_FATURADA'].sum() > 0 else 0,
                df_resultado['QTD_NOTAS_FATURADAS'].sum(),
                df_resultado['QTD_NOTAS_DEVOLVIDAS'].sum()
            ]
        }
        return pd.DataFrame(resumo_data)
    
    def _criar_analise_estabelecimento(self, df_resultado):
        """Cria análise por estabelecimento"""
        estabelecimentos = ['R351', 'R352']
        analise_data = []
        
        for estabelecimento in estabelecimentos:
            df_filtrado = df_resultado[df_resultado['COD_ESTABELECIMENTO'] == estabelecimento]
            if len(df_filtrado) > 0:
                analise_data.append({
                    'Estabelecimento': estabelecimento,
                    'Total Vendedores': df_filtrado['VENDEDOR'].nunique(),
                    'Valor Faturado (R$)': df_filtrado['VALOR_FATURAMENTO'].sum(),
                    'Valor Devolvido (R$)': df_filtrado['VALOR_DEVOLUCAO'].sum(),
                    'Percentual Devolução (%)': (df_filtrado['VALOR_DEVOLUCAO'].sum() / df_filtrado['VALOR_FATURAMENTO'].sum() * 100) if df_filtrado['VALOR_FATURAMENTO'].sum() > 0 else 0,
                    'Média Índice Devolução': df_filtrado['INDICE_DEVOLUCAO'].mean(),
                    'Média % Devolução': df_filtrado['PERCENTUAL_DEVOLUCAO'].mean(),
                    'Quantidade Faturada': df_filtrado['TOTAL_QTD_FATURADA'].sum(),
                    'Quantidade Devolvida': df_filtrado['TOTAL_QTD_DEVOLVIDA'].sum(),
                    'Notas Faturadas': df_filtrado['QTD_NOTAS_FATURADAS'].sum(),
                    'Notas Devolvidas': df_filtrado['QTD_NOTAS_DEVOLVIDAS'].sum()
                })
        
        # Adicionar linha de total
        analise_data.append({
            'Estabelecimento': 'TOTAL',
            'Total Vendedores': df_resultado['VENDEDOR'].nunique(),
            'Valor Faturado (R$)': df_resultado['VALOR_FATURAMENTO'].sum(),
            'Valor Devolvido (R$)': df_resultado['VALOR_DEVOLUCAO'].sum(),
            'Percentual Devolução (%)': (df_resultado['VALOR_DEVOLUCAO'].sum() / df_resultado['VALOR_FATURAMENTO'].sum() * 100) if df_resultado['VALOR_FATURAMENTO'].sum() > 0 else 0,
            'Média Índice Devolução': df_resultado['INDICE_DEVOLUCAO'].mean(),
            'Média % Devolução': df_resultado['PERCENTUAL_DEVOLUCAO'].mean(),
            'Quantidade Faturada': df_resultado['TOTAL_QTD_FATURADA'].sum(),
            'Quantidade Devolvida': df_resultado['TOTAL_QTD_DEVOLVIDA'].sum(),
            'Notas Faturadas': df_resultado['QTD_NOTAS_FATURADAS'].sum(),
            'Notas Devolvidas': df_resultado['QTD_NOTAS_DEVOLVIDAS'].sum()
        })
        
        return pd.DataFrame(analise_data)
    
    def gerar_relatorio_resumo(self, df_resultado):
        """Gera um resumo estatístico do índice de devolução"""
        print("\n" + "="*80)
        print("RESUMO DO ÍNDICE DE DEVOLUÇÃO POR VENDEDOR")
        print("="*80)
        
        # Estatísticas gerais
        print(f"\nTotal de Vendedores: {len(df_resultado)}")
        print(f"Vendedores Únicos: {df_resultado['VENDEDOR'].nunique()}")
        print(f"Média do Índice de Devolução: {df_resultado['INDICE_DEVOLUCAO'].mean():.4f}")
        print(f"Média do Percentual de Devolução: {df_resultado['PERCENTUAL_DEVOLUCAO'].mean():.2f}%")
        print(f"Mediana do Índice de Devolução: {df_resultado['INDICE_DEVOLUCAO'].median():.4f}")
        print(f"Máximo Índice de Devolução: {df_resultado['INDICE_DEVOLUCAO'].max():.4f}")
        print(f"Mínimo Índice de Devolução: {df_resultado['INDICE_DEVOLUCAO'].min():.4f}")
        
        # Valor total faturado e devolvido
        total_faturado = df_resultado['VALOR_FATURAMENTO'].sum()
        total_devolvido = df_resultado['VALOR_DEVOLUCAO'].sum()
        percentual_total = (total_devolvido / total_faturado * 100) if total_faturado > 0 else 0
        
        print(f"\nValor Total Faturado: R$ {total_faturado:,.2f}")
        print(f"Valor Total Devolvido: R$ {total_devolvido:,.2f}")
        print(f"Percentual Total de Devolução: {percentual_total:.2f}%")
        
        # Quantidade total
        total_qtd_faturada = df_resultado['TOTAL_QTD_FATURADA'].sum()
        total_qtd_devolvida = df_resultado['TOTAL_QTD_DEVOLVIDA'].sum()
        percentual_qtd = (total_qtd_devolvida / total_qtd_faturada * 100) if total_qtd_faturada > 0 else 0
        
        print(f"\nQuantidade Total Faturada: {total_qtd_faturada:,.0f}")
        print(f"Quantidade Total Devolvida: {total_qtd_devolvida:,.0f}")
        print(f"Percentual Quantidade Devolvida: {percentual_qtd:.2f}%")
        
        # Top 10 maiores índices
        print(f"\nTOP 10 MAIORES ÍNDICES DE DEVOLUÇÃO:")
        print("-"*80)
        top_10_cols = ['RANKING', 'COD_ESTABELECIMENTO', 'VENDEDOR', 
                      'VALOR_FATURAMENTO', 'VALOR_DEVOLUCAO', 'PERCENTUAL_DEVOLUCAO']
        
        # Verificar se todas as colunas existem
        top_10_cols = [col for col in top_10_cols if col in df_resultado.columns]
        
        top_10 = df_resultado.head(10)[top_10_cols]
        print(top_10.to_string(index=False))
        
        # Análise por estabelecimento
        print(f"\nANÁLISE POR ESTABELECIMENTO:")
        print("-"*80)
        for estabelecimento in ['R351', 'R352']:
            df_filtrado = df_resultado[df_resultado['COD_ESTABELECIMENTO'] == estabelecimento]
            if len(df_filtrado) > 0:
                print(f"\n{estabelecimento}:")
                print(f"  Vendedores: {len(df_filtrado)}")
                print(f"  Média Percentual Devolução: {df_filtrado['PERCENTUAL_DEVOLUCAO'].mean():.2f}%")
                print(f"  Valor Total Faturado: R$ {df_filtrado['VALOR_FATURAMENTO'].sum():,.2f}")
                print(f"  Valor Total Devolvido: R$ {df_filtrado['VALOR_DEVOLUCAO'].sum():,.2f}")
                print(f"  Quantidade Faturada: {df_filtrado['TOTAL_QTD_FATURADA'].sum():,.0f}")
                print(f"  Quantidade Devolvida: {df_filtrado['TOTAL_QTD_DEVOLVIDA'].sum():,.0f}")
    
    def executar_analise_completa(self, data_inicio_faturamento='2025-07-01', 
                                 data_fim_faturamento='2025-12-31',
                                 data_inicio_devolucao='2025-07-01',
                                 data_fim_devolucao='2026-12-31',
                                 exportar=True):
        """Executa toda a análise completa"""
        
        print(f"\nDiretório de saída configurado: {self.caminho_base}")
        
        # Conectar ao banco
        if not self.conectar_banco():
            return
        
        try:
            # Obter dados
            df_devolucao = self.obter_dados_devolucao(data_inicio_devolucao, data_fim_devolucao)
            df_faturamento = self.obter_dados_faturamento(data_inicio_faturamento, data_fim_faturamento)
            
            if df_devolucao is not None and df_faturamento is not None:
                print(f"\nDados de devolução obtidos: {len(df_devolucao)} registros")
                print(f"Dados de faturamento obtidos: {len(df_faturamento)} registros")
                
                # Verificar se há dados
                if len(df_devolucao) == 0:
                    print("AVISO: Nenhum dado de devolução encontrado para o período especificado.")
                if len(df_faturamento) == 0:
                    print("AVISO: Nenhum dado de faturamento encontrado para o período especificado.")
                
                # Calcular índice
                df_resultado = self.calcular_indice_devolucao(df_devolucao, df_faturamento)
                
                if df_resultado is not None and len(df_resultado) > 0:
                    print(f"\nCálculo do índice concluído: {len(df_resultado)} registros processados")
                    
                    # Exibir primeiras linhas
                    print("\nPrimeiras linhas do resultado:")
                    print(df_resultado.head().to_string(index=False))
                    
                    # Gerar relatório resumo
                    self.gerar_relatorio_resumo(df_resultado)
                    
                    # Exportar resultados
                    if exportar:
                        print("\n" + "="*80)
                        print("EXPORTANDO RESULTADOS PARA EXCEL...")
                        print("="*80)
                        arquivo = self.exportar_resultados(df_resultado, formato='excel')
                        if arquivo:
                            print(f"\nArquivo Excel criado com sucesso!")
                            print(f"Local: {arquivo}")
                            
                            # Abrir o arquivo (opcional - descomente se quiser)
                            # os.startfile(arquivo)
                        
                    return df_resultado
                else:
                    print("Erro ao calcular índice de devolução ou nenhum dado retornado")
                    return None
            else:
                print("Erro ao obter dados do banco")
                return None
                
        except Exception as e:
            print(f"Erro durante a execução: {e}")
            import traceback
            traceback.print_exc()
            return None
        finally:
            # Fechar conexão
            if self.conn:
                self.conn.close()
                print("\nConexão com o banco de dados fechada.")

def main():
    """
    Função principal para executar o script
    """
    print("="*80)
    print("CÁLCULO DO ÍNDICE DE DEVOLUÇÃO POR VENDEDOR")
    print("="*80)
    
    # Criar instância com as credenciais fornecidas
    analise = CalculoIndiceDevolucao()
    
    # Datas personalizáveis (opcional)
    datas = {
        'data_inicio_faturamento': '2025-07-01',
        'data_fim_faturamento': '2025-12-31',
        'data_inicio_devolucao': '2025-07-01',
        'data_fim_devolucao': '2026-12-31'
    }
    
    print(f"\nPeríodo de análise:")
    print(f"Faturamento: {datas['data_inicio_faturamento']} até {datas['data_fim_faturamento']}")
    print(f"Devolução: {datas['data_inicio_devolucao']} até {datas['data_fim_devolucao']}")
    
    # Perguntar se deseja usar datas diferentes
    alterar_datas = input("\nDeseja alterar as datas? (s/n): ").strip().lower()
    
    if alterar_datas == 's':
        datas['data_inicio_faturamento'] = input("Data início faturamento (YYYY-MM-DD): ").strip()
        datas['data_fim_faturamento'] = input("Data fim faturamento (YYYY-MM-DD): ").strip()
        datas['data_inicio_devolucao'] = input("Data início devolução (YYYY-MM-DD): ").strip()
        datas['data_fim_devolucao'] = input("Data fim devolução (YYYY-MM-DD): ").strip()
    
    # Executar análise completa
    resultados = analise.executar_analise_completa(**datas, exportar=True)
    
    if resultados is not None:
        print("\n" + "="*80)
        print("ANÁLISE CONCLUÍDA COM SUCESSO!")
        print("="*80)
        print(f"Total de vendedores processados: {len(resultados)}")
        
        # Mostrar onde o arquivo foi salvo
        print(f"\nArquivo Excel salvo em: {analise.caminho_base}")
        print("O arquivo contém múltiplas planilhas com análises detalhadas.")
    else:
        print("\n" + "="*80)
        print("FALHA NA ANÁLISE")
        print("="*80)

# Versão simplificada para uso rápido
def versao_rapida():
    """
    Versão rápida sem interação
    """
    print("Executando versão rápida com datas padrão...")
    
    analise = CalculoIndiceDevolucao()
    
    # Usar datas padrão
    resultados = analise.executar_analise_completa(
        data_inicio_faturamento='2025-07-01',
        data_fim_faturamento='2025-12-31',
        data_inicio_devolucao='2025-07-01',
        data_fim_devolucao='2026-12-31',
        exportar=True
    )
    
    return resultados

# Executar o script
if __name__ == "__main__":
    print("="*80)
    print("CÁLCULO DO ÍNDICE DE DEVOLUÇÃO POR VENDEDOR")
    print("="*80)
    
    # Escolher modo de execução
    print("\nModo de execução:")
    print("1 - Versão interativa (permite alterar datas)")
    print("2 - Versão rápida (usa datas padrão)")
    
    try:
        modo = input("Escolha (1/2): ").strip()
        
        if modo == '1':
            main()
        elif modo == '2':
            versao_rapida()
        else:
            print("Opção inválida. Executando versão interativa...")
            main()
    except KeyboardInterrupt:
        print("\n\nExecução interrompida pelo usuário.")
    except Exception as e:
        print(f"\nErro inesperado: {e}")
    finally:
        input("\nPressione Enter para sair...")
