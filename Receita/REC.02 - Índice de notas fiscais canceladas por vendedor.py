import pyodbc
import pandas as pd
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')

class CalculoIndiceCancelamento:
    def __init__(self):
        """
        Inicializa com as credenciais fornecidas
        """
        self.server = 'DCMDWF01A.MOURA.INT'
        self.database = 'ax'
        self.username = 'uAuditoria'
        self.password = '@ud!t0$!@202&22'
        self.driver = '{SQL Server}'
        self.conn = None
        
        # Definir o caminho específico para salvar
        self.caminho_base = r'C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2026\03. Automações\Validações'
        
        # Parâmetros padrão
        self.cod_estabelecimento = 'R281'
        self.data_inicio = '2025-01-01'
        self.data_fim = '2025-12-31'
    
    def definir_parametros(self, cod_estabelecimento=None, data_inicio=None, data_fim=None):
        """
        Define os parâmetros para a análise
        """
        if cod_estabelecimento:
            self.cod_estabelecimento = cod_estabelecimento
        if data_inicio:
            self.data_inicio = data_inicio
        if data_fim:
            self.data_fim = data_fim
            
        print(f"✅ Parâmetros definidos:")
        print(f"   Estabelecimento: {self.cod_estabelecimento}")
        print(f"   Data início: {self.data_inicio}")
        print(f"   Data fim: {self.data_fim}")
    
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
            print("✅ Conexão estabelecida com sucesso!")
            return True
        except Exception as e:
            print(f"❌ Erro ao conectar ao banco de dados: {e}")
            return False
    
    def executar_query(self, query):
        """Executa uma query SQL e retorna um DataFrame"""
        try:
            df = pd.read_sql_query(query, self.conn)
            return df
        except Exception as e:
            print(f"❌ Erro ao executar query: {e}")
            return None
    
    def obter_dados_cancelamentos(self):
        """Obtém os dados de notas canceladas por vendedor"""
        query_cancelamentos = f"""
        SELECT
            COD_ESTABELECIMENTO,
            VENDEDOR,
            SUM(QUANTIDADE) AS TOTAL_QTD_CANCELADA,
            SUM(VALOR) AS VALOR_CANCELAMENTO,
            COUNT(*) AS QTD_NOTAS_CANCELADAS
        FROM 
            VW_AUDIT_RM_ORDENS_VENDA
        WHERE
            COD_ESTABELECIMENTO = '{self.cod_estabelecimento}'
            AND DATA_NOTA_FISCAL BETWEEN '{self.data_inicio}' AND '{self.data_fim}' 
            AND PARA_FATURAMENTO = 'Sim'
            AND NUM_NOTA_FISCAL LIKE '%CAN%'
        GROUP BY
            COD_ESTABELECIMENTO,
            VENDEDOR
        HAVING 
            COUNT(*) <> 0
        ORDER BY
            COD_ESTABELECIMENTO,
            VENDEDOR,
            VALOR_CANCELAMENTO ASC
        """
        
        print(f"📊 Executando query de notas canceladas para estabelecimento {self.cod_estabelecimento}...")
        print(f"   Período: {self.data_inicio} a {self.data_fim}")
        df_cancelamentos = self.executar_query(query_cancelamentos)
        return df_cancelamentos
    
    def obter_dados_faturamento(self):
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
            COD_ESTABELECIMENTO = '{self.cod_estabelecimento}'
            AND DATA_NOTA_FISCAL BETWEEN '{self.data_inicio}' AND '{self.data_fim}'  
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
        
        print(f"📈 Executando query de faturamento para estabelecimento {self.cod_estabelecimento}...")
        print(f"   Período: {self.data_inicio} a {self.data_fim}")
        df_faturamento = self.executar_query(query_faturamento)
        return df_faturamento
    
    def calcular_indice_cancelamento(self, df_cancelamentos, df_faturamento):
        """Calcula o índice de cancelamento: VALOR_CANCELADO / VALOR_FATURADO"""
        if df_cancelamentos is None or df_faturamento is None:
            print("❌ Erro: Dados não disponíveis para cálculo")
            return None
        
        # Criar chave única para merge
        df_faturamento['chave'] = df_faturamento['COD_ESTABELECIMENTO'] + '_' + df_faturamento['VENDEDOR'].astype(str)
        df_cancelamentos['chave'] = df_cancelamentos['COD_ESTABELECIMENTO'] + '_' + df_cancelamentos['VENDEDOR'].astype(str)
        
        # Realizar merge dos dataframes (LEFT JOIN para manter todos os vendedores com faturamento)
        df_consolidado = pd.merge(
            df_faturamento,
            df_cancelamentos[['chave', 'VALOR_CANCELAMENTO', 'TOTAL_QTD_CANCELADA', 'QTD_NOTAS_CANCELADAS']],
            on='chave',
            how='left'
        )
        
        # Preencher valores NaN com 0 (vendedores sem cancelamentos)
        df_consolidado['VALOR_CANCELAMENTO'] = df_consolidado['VALOR_CANCELAMENTO'].fillna(0)
        df_consolidado['TOTAL_QTD_CANCELADA'] = df_consolidado['TOTAL_QTD_CANCELADA'].fillna(0)
        df_consolidado['QTD_NOTAS_CANCELADAS'] = df_consolidado['QTD_NOTAS_CANCELADAS'].fillna(0)
        
        # Calcular índice de cancelamento
        df_consolidado['INDICE_CANCELAMENTO'] = df_consolidado.apply(
            lambda row: row['VALOR_CANCELAMENTO'] / row['VALOR_FATURAMENTO'] 
            if row['VALOR_FATURAMENTO'] > 0 else 0,
            axis=1
        )
        
        # Calcular percentual de cancelamento
        df_consolidado['PERCENTUAL_CANCELAMENTO'] = df_consolidado['INDICE_CANCELAMENTO'] * 100
        
        # Calcular percentual de quantidade cancelada
        df_consolidado['PERCENTUAL_QTD_CANCELADA'] = df_consolidado.apply(
            lambda row: (row['TOTAL_QTD_CANCELADA'] / row['TOTAL_QTD_FATURADA'] * 100)
            if row['TOTAL_QTD_FATURADA'] > 0 else 0,
            axis=1
        )
        
        # Calcular ticket médio
        df_consolidado['TICKET_MEDIO_FATURAMENTO'] = df_consolidado.apply(
            lambda row: row['VALOR_FATURAMENTO'] / row['QTD_NOTAS_FATURADAS']
            if row['QTD_NOTAS_FATURADAS'] > 0 else 0,
            axis=1
        )
        
        df_consolidado['VALOR_MEDIO_CANCELAMENTO'] = df_consolidado.apply(
            lambda row: row['VALOR_CANCELAMENTO'] / row['QTD_NOTAS_CANCELADAS']
            if row['QTD_NOTAS_CANCELADAS'] > 0 else 0,
            axis=1
        )
        
        # Classificar por risco de cancelamento
        df_consolidado['CLASSIFICACAO_RISCO'] = df_consolidado.apply(
            lambda row: self._classificar_risco(row['PERCENTUAL_CANCELAMENTO']),
            axis=1
        )
        
        # Ordenar por maior índice de cancelamento
        df_consolidado = df_consolidado.sort_values(['INDICE_CANCELAMENTO', 'VALOR_CANCELAMENTO'], ascending=[False, False])
        
        # Adicionar ranking
        df_consolidado['RANKING'] = range(1, len(df_consolidado) + 1)
        
        # Remover coluna auxiliar
        df_consolidado = df_consolidado.drop('chave', axis=1)
        
        # Reorganizar colunas
        colunas = [
            'RANKING', 'COD_ESTABELECIMENTO', 'VENDEDOR', 'CLASSIFICACAO_RISCO',
            'VALOR_FATURAMENTO', 'QTD_NOTAS_FATURADAS', 'TOTAL_QTD_FATURADA', 'TICKET_MEDIO_FATURAMENTO',
            'VALOR_CANCELAMENTO', 'QTD_NOTAS_CANCELADAS', 'TOTAL_QTD_CANCELADA', 'VALOR_MEDIO_CANCELAMENTO',
            'INDICE_CANCELAMENTO', 'PERCENTUAL_CANCELAMENTO', 'PERCENTUAL_QTD_CANCELADA'
        ]
        
        # Manter apenas as colunas que existem
        colunas = [col for col in colunas if col in df_consolidado.columns]
        df_consolidado = df_consolidado[colunas]
        
        return df_consolidado
    
    def _classificar_risco(self, percentual):
        """Classifica o vendedor pelo percentual de cancelamento"""
        if percentual == 0:
            return 'SEM CANCELAMENTO'
        elif percentual > 10:
            return 'ALTO RISCO'
        elif percentual > 5:
            return 'MÉDIO RISCO'
        else:
            return 'BAIXO RISCO'
    
    def exportar_resultados(self, df_resultado, formato='excel'):
        """Exporta os resultados para Excel ou CSV no caminho especificado"""
        try:
            # Verificar se o diretório existe, se não, criar
            if not os.path.exists(self.caminho_base):
                print(f"📁 Criando diretório: {self.caminho_base}")
                os.makedirs(self.caminho_base)
            
            # Criar nome do arquivo com estabelecimento, período, data e hora
            estabelecimento_codigo = self.cod_estabelecimento.replace('/', '_')
            data_inicio_formatada = self.data_inicio.replace('-', '')
            data_fim_formatada = self.data_fim.replace('-', '')
            data_atual = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            if formato.lower() == 'excel':
                nome_arquivo = f'indice_cancelamento_{estabelecimento_codigo}_{data_inicio_formatada}_a_{data_fim_formatada}_{data_atual}.xlsx'
                caminho_completo = os.path.join(self.caminho_base, nome_arquivo)
                
                # Garantir que os nomes das colunas sejam strings válidas para Excel
                df_resultado.columns = [str(col).replace('/', '_').replace('[', '').replace(']', '') 
                                       for col in df_resultado.columns]
                
                # Exportar para Excel com APENAS a planilha principal
                with pd.ExcelWriter(caminho_completo, engine='openpyxl') as writer:
                    # Apenas a planilha principal com todos os dados
                    df_resultado.to_excel(writer, sheet_name='Indice_Cancelamento', index=False)
                
                print(f"💾 Resultados exportados para: {caminho_completo}")
                
                # Também criar um arquivo CSV
                nome_csv = f'indice_cancelamento_{estabelecimento_codigo}_{data_inicio_formatada}_a_{data_fim_formatada}_{data_atual}.csv'
                caminho_csv = os.path.join(self.caminho_base, nome_csv)
                df_resultado.to_csv(caminho_csv, index=False, sep=';', decimal=',', encoding='utf-8')
                print(f"📄 Arquivo CSV também criado em: {caminho_csv}")
                
            else:
                nome_arquivo = f'indice_cancelamento_{estabelecimento_codigo}_{data_inicio_formatada}_a_{data_fim_formatada}_{data_atual}.csv'
                caminho_completo = os.path.join(self.caminho_base, nome_arquivo)
                df_resultado.to_csv(caminho_completo, index=False, sep=';', decimal=',', encoding='utf-8')
                print(f"💾 Resultados exportados para: {caminho_completo}")
            
            return caminho_completo
            
        except Exception as e:
            print(f"❌ Erro ao exportar resultados: {e}")
            # Tentar salvar no diretório atual como fallback
            try:
                data_atual = datetime.now().strftime('%Y%m%d_%H%M%S')
                nome_arquivo = f'indice_cancelamento_vendedores_{data_atual}.xlsx'
                df_resultado.to_excel(nome_arquivo, index=False)
                print(f"💾 Resultados exportados para o diretório atual: {nome_arquivo}")
                return nome_arquivo
            except:
                print("❌ Não foi possível exportar os resultados.")
                return None
    
    def gerar_relatorio_resumo(self, df_resultado):
        """Gera um resumo estatístico do índice de cancelamento no console"""
        print("\n" + "="*80)
        print(f"RESUMO DO ÍNDICE DE CANCELAMENTO DE NOTAS POR VENDEDOR")
        print(f"Estabelecimento: {self.cod_estabelecimento}")
        print(f"Período: {self.data_inicio} a {self.data_fim}")
        print("="*80)
        
        # Estatísticas gerais
        print(f"\n📊 Total de Vendedores: {len(df_resultado)}")
        print(f"📊 Vendedores com Cancelamento: {(df_resultado['VALOR_CANCELAMENTO'] > 0).sum()}")
        print(f"📊 Vendedores sem Cancelamento: {(df_resultado['VALOR_CANCELAMENTO'] == 0).sum()}")
        print(f"📊 Média do Índice de Cancelamento: {df_resultado['INDICE_CANCELAMENTO'].mean():.4f}")
        print(f"📊 Média do Percentual de Cancelamento: {df_resultado['PERCENTUAL_CANCELAMENTO'].mean():.2f}%")
        
        # Valor total faturado e cancelado
        total_faturado = df_resultado['VALOR_FATURAMENTO'].sum()
        total_cancelado = df_resultado['VALOR_CANCELAMENTO'].sum()
        percentual_total = (total_cancelado / total_faturado * 100) if total_faturado > 0 else 0
        
        print(f"\n💰 Valor Total Faturado: R$ {total_faturado:,.2f}")
        print(f"💰 Valor Total Cancelado: R$ {total_cancelado:,.2f}")
        print(f"💰 Percentual Total de Cancelamento: {percentual_total:.2f}%")
        
        # Quantidade total
        total_qtd_faturada = df_resultado['TOTAL_QTD_FATURADA'].sum()
        total_qtd_cancelada = df_resultado['TOTAL_QTD_CANCELADA'].sum()
        percentual_qtd = (total_qtd_cancelada / total_qtd_faturada * 100) if total_qtd_faturada > 0 else 0
        
        print(f"\n📦 Quantidade Total Faturada: {total_qtd_faturada:,.0f}")
        print(f"📦 Quantidade Total Cancelada: {total_qtd_cancelada:,.0f}")
        print(f"📦 Percentual Quantidade Cancelada: {percentual_qtd:.2f}%")
        
        # Top 10 maiores índices
        print(f"\n🏆 TOP 10 MAIORES ÍNDICES DE CANCELAMENTO:")
        print("-"*80)
        top_10_cols = ['RANKING', 'COD_ESTABELECIMENTO', 'VENDEDOR', 'CLASSIFICACAO_RISCO',
                      'VALOR_FATURAMENTO', 'VALOR_CANCELAMENTO', 'PERCENTUAL_CANCELAMENTO']
        
        top_10_cols = [col for col in top_10_cols if col in df_resultado.columns]
        top_10 = df_resultado.head(10)[top_10_cols]
        print(top_10.to_string(index=False))
        
        # Análise por classificação de risco (apenas no console)
        print(f"\n⚠️  ANÁLISE POR CLASSIFICAÇÃO DE RISCO:")
        print("-"*80)
        for classificacao in ['ALTO RISCO', 'MÉDIO RISCO', 'BAIXO RISCO', 'SEM CANCELAMENTO']:
            df_filtrado = df_resultado[df_resultado['CLASSIFICACAO_RISCO'] == classificacao]
            if len(df_filtrado) > 0:
                print(f"\n{classificacao}:")
                print(f"  Vendedores: {len(df_filtrado)} ({len(df_filtrado)/len(df_resultado)*100:.1f}%)")
                print(f"  Valor Faturado: R$ {df_filtrado['VALOR_FATURAMENTO'].sum():,.2f}")
                print(f"  Valor Cancelado: R$ {df_filtrado['VALOR_CANCELAMENTO'].sum():,.2f}")
        
        # Análise por estabelecimento atual (apenas no console)
        print(f"\n🏢 ANÁLISE POR ESTABELECIMENTO:")
        print("-"*80)
        df_filtrado = df_resultado[df_resultado['COD_ESTABELECIMENTO'] == self.cod_estabelecimento]
        if len(df_filtrado) > 0:
            print(f"\n{self.cod_estabelecimento}:")
            print(f"  Vendedores: {len(df_filtrado)}")
            print(f"  Média % Cancelamento: {df_filtrado['PERCENTUAL_CANCELAMENTO'].mean():.2f}%")
            print(f"  Valor Faturado: R$ {df_filtrado['VALOR_FATURAMENTO'].sum():,.2f}")
            print(f"  Valor Cancelado: R$ {df_filtrado['VALOR_CANCELAMENTO'].sum():,.2f}")
    
    def executar_analise_completa(self, cod_estabelecimento=None, data_inicio=None, 
                                 data_fim=None, exportar=True):
        """Executa toda a análise completa com parâmetros personalizáveis"""
        
        # Definir parâmetros se fornecidos
        if cod_estabelecimento:
            self.cod_estabelecimento = cod_estabelecimento
        if data_inicio:
            self.data_inicio = data_inicio
        if data_fim:
            self.data_fim = data_fim
        
        print(f"\n📂 Diretório de saída configurado: {self.caminho_base}")
        print(f"🏢 Estabelecimento: {self.cod_estabelecimento}")
        print(f"📅 Período: {self.data_inicio} a {self.data_fim}")
        
        # Conectar ao banco
        if not self.conectar_banco():
            return
        
        try:
            # Obter dados
            df_cancelamentos = self.obter_dados_cancelamentos()
            df_faturamento = self.obter_dados_faturamento()
            
            if df_cancelamentos is not None and df_faturamento is not None:
                print(f"\n📊 Dados de cancelamentos obtidos: {len(df_cancelamentos)} registros")
                print(f"📈 Dados de faturamento obtidos: {len(df_faturamento)} registros")
                
                # Verificar se há dados
                if len(df_cancelamentos) == 0:
                    print("⚠️  AVISO: Nenhum dado de cancelamento encontrado para o período especificado.")
                if len(df_faturamento) == 0:
                    print("⚠️  AVISO: Nenhum dado de faturamento encontrado para o período especificado.")
                
                # Calcular índice
                df_resultado = self.calcular_indice_cancelamento(df_cancelamentos, df_faturamento)
                
                if df_resultado is not None and len(df_resultado) > 0:
                    print(f"\n✅ Cálculo do índice concluído: {len(df_resultado)} registros processados")
                    
                    # Exibir primeiras linhas
                    print("\n📋 Primeiras linhas do resultado:")
                    print(df_resultado.head().to_string(index=False))
                    
                    # Gerar relatório resumo (apenas no console)
                    self.gerar_relatorio_resumo(df_resultado)
                    
                    # Exportar resultados
                    if exportar:
                        print("\n" + "="*80)
                        print("💾 EXPORTANDO RESULTADOS PARA EXCEL...")
                        print("="*80)
                        arquivo = self.exportar_resultados(df_resultado, formato='excel')
                        if arquivo:
                            print(f"\n✅ Arquivo Excel criado com sucesso!")
                            print(f"📍 Local: {arquivo}")
                            print("📑 O arquivo contém apenas a planilha 'Indice_Cancelamento' com todos os dados")
                        
                    return df_resultado
                else:
                    print("❌ Erro ao calcular índice de cancelamento ou nenhum dado retornado")
                    return None
            else:
                print("❌ Erro ao obter dados do banco")
                return None
                
        except Exception as e:
            print(f"❌ Erro durante a execução: {e}")
            import traceback
            traceback.print_exc()
            return None
        finally:
            # Fechar conexão
            if self.conn:
                self.conn.close()
                print("\n🔒 Conexão com o banco de dados fechada.")

def main():
    """
    Função principal para executar o script
    """
    print("="*80)
    print("📊 CÁLCULO DO ÍNDICE DE CANCELAMENTO DE NOTAS POR VENDEDOR")
    print("="*80)
    
    # Criar instância com as credenciais fornecidas
    analise = CalculoIndiceCancelamento()
    
    # Parâmetros personalizáveis
    parametros = {
        'cod_estabelecimento': 'R281',
        'data_inicio': '2025-01-01',
        'data_fim': '2025-12-31'
    }
    
    print(f"\n📋 Parâmetros atuais:")
    print(f"   Estabelecimento: {parametros['cod_estabelecimento']}")
    print(f"   Data início: {parametros['data_inicio']}")
    print(f"   Data fim: {parametros['data_fim']}")
    
    # Perguntar se deseja alterar os parâmetros
    alterar_parametros = input("\n📝 Deseja alterar os parâmetros? (s/n): ").strip().lower()
    
    if alterar_parametros == 's':
        print("\n📝 Insira os novos parâmetros (deixe em branco para manter o atual):")
        
        novo_estabelecimento = input(f"Estabelecimento [{parametros['cod_estabelecimento']}]: ").strip()
        if novo_estabelecimento:
            parametros['cod_estabelecimento'] = novo_estabelecimento
        
        nova_data_inicio = input(f"Data início [{parametros['data_inicio']}]: ").strip()
        if nova_data_inicio:
            parametros['data_inicio'] = nova_data_inicio
        
        nova_data_fim = input(f"Data fim [{parametros['data_fim']}]: ").strip()
        if nova_data_fim:
            parametros['data_fim'] = nova_data_fim
    
    # Executar análise completa
    resultados = analise.executar_analise_completa(
        cod_estabelecimento=parametros['cod_estabelecimento'],
        data_inicio=parametros['data_inicio'],
        data_fim=parametros['data_fim'],
        exportar=True
    )
    
    if resultados is not None:
        print("\n" + "="*80)
        print("✅ ANÁLISE CONCLUÍDA COM SUCESSO!")
        print("="*80)
        print(f"🏢 Estabelecimento analisado: {analise.cod_estabelecimento}")
        print(f"📅 Período: {analise.data_inicio} a {analise.data_fim}")
        print(f"📊 Total de vendedores processados: {len(resultados)}")
        print(f"📂 Arquivo Excel salvo em: {analise.caminho_base}")
        print("📑 O arquivo contém apenas a planilha 'Indice_Cancelamento' com todos os dados")
    else:
        print("\n" + "="*80)
        print("❌ FALHA NA ANÁLISE")
        print("="*80)

# Versão simplificada para uso rápido
def versao_rapida():
    """
    Versão rápida sem interação
    """
    print("⚡ Executando versão rápida com parâmetros padrão...")
    
    analise = CalculoIndiceCancelamento()
    
    # Usar parâmetros padrão
    resultados = analise.executar_analise_completa(
        cod_estabelecimento='R281',
        data_inicio='2025-01-01',
        data_fim='2025-12-31',
        exportar=True
    )
    
    return resultados

# Nova função para analisar múltiplos estabelecimentos
def analisar_multiplos_estabelecimentos():
    """
    Analisa múltiplos estabelecimentos em sequência
    """
    print("="*80)
    print("🏢 ANÁLISE DE MÚLTIPLOS ESTABELECIMENTOS")
    print("="*80)
    
    # Solicitar estabelecimentos
    estabelecimentos_input = input("Digite os códigos dos estabelecimentos (separados por vírgula): ").strip()
    estabelecimentos = [e.strip() for e in estabelecimentos_input.split(',') if e.strip()]
    
    if not estabelecimentos:
        print("❌ Nenhum estabelecimento informado.")
        return
    
    print(f"\n📋 Estabelecimentos a serem analisados: {', '.join(estabelecimentos)}")
    
    # Solicitar período
    data_inicio = input(f"Data início [2025-01-01]: ").strip() or '2025-01-01'
    data_fim = input(f"Data fim [2025-12-31]: ").strip() or '2025-12-31'
    
    print(f"\n📅 Período para todos os estabelecimentos: {data_inicio} a {data_fim}")
    
    confirmar = input("\n📝 Confirmar análise? (s/n): ").strip().lower()
    
    if confirmar != 's':
        print("❌ Análise cancelada.")
        return
    
    resultados_completos = []
    
    for estabelecimento in estabelecimentos:
        print(f"\n{'='*60}")
        print(f"📊 ANALISANDO ESTABELECIMENTO: {estabelecimento}")
        print(f"{'='*60}")
        
        analise = CalculoIndiceCancelamento()
        
        # Executar análise para este estabelecimento
        resultado = analise.executar_analise_completa(
            cod_estabelecimento=estabelecimento,
            data_inicio=data_inicio,
            data_fim=data_fim,
            exportar=True
        )
        
        if resultado is not None:
            resultados_completos.append((estabelecimento, resultado))
            print(f"✅ Estabelecimento {estabelecimento} analisado com sucesso!")
        else:
            print(f"❌ Falha na análise do estabelecimento {estabelecimento}")
    
    print(f"\n{'='*80}")
    print("📊 RESUMO DA ANÁLISE DE MÚLTIPLOS ESTABELECIMENTOS")
    print(f"{'='*80}")
    
    for estabelecimento, resultado in resultados_completos:
        if resultado is not None:
            print(f"🏢 {estabelecimento}: {len(resultado)} vendedores analisados ✓")
    
    return resultados_completos

# Executar o script
if __name__ == "__main__":
    print("="*80)
    print("📊 CÁLCULO DO ÍNDICE DE CANCELAMENTO DE NOTAS POR VENDEDOR")
    print("="*80)
    
    # Escolher modo de execução
    print("\n🎯 Modo de execução:")
    print("1 - Versão interativa (permite alterar estabelecimento e datas)")
    print("2 - Versão rápida (usa parâmetros padrão)")
    print("3 - Análise de múltiplos estabelecimentos")
    
    try:
        modo = input("\nEscolha (1/2/3): ").strip()
        
        if modo == '1':
            main()
        elif modo == '2':
            versao_rapida()
        elif modo == '3':
            analisar_multiplos_estabelecimentos()
        else:
            print("⚠️  Opção inválida. Executando versão interativa...")
            main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Execução interrompida pelo usuário.")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
    finally:
        input("\n👋 Pressione Enter para sair...")
