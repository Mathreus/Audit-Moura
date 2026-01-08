# OBSERVAÇÕES SOBRE O CÓDIGO!!!

# 1) É importante salvar o arquivo TXT da folha em uma pasta no Drive;
# 2) O caminho dessa pasta no DRIVE será inserido dentro do código, na linha 22;
# 3) Quando salvar o arquivo TXT, por gentileza, não utilizar caracteres especiaos como acentos, @, #,$, %, de preferência salvem como exemplo "FOPAG_BAURU.txt"
# 4) A planilha será salva no mesmo caminho criado

# Processar_folha.py
import re
import pandas as pd
import os
from datetime import datetime
from collections import defaultdict

# Configurações de exibição do pandas
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.colheader_justify', 'center')

# Caminhos dos arquivos
CAMINHO_BASE = r"C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Automações\Folha\Bauru"
NOME_ARQUIVO_TXT = "Fopag_Bauru.txt"
CAMINHO_ARQUIVO_TXT = os.path.join(CAMINHO_BASE, NOME_ARQUIVO_TXT)

# Lista de códigos de proventos e descontos (será preenchida dinamicamente)
PROVENTOS_CODIGOS = set()
DESCONTOS_CODIGOS = set()

def identificar_proventos_descontos_no_arquivo(filepath):
    """
    Identifica todos os tipos de proventos e descontos presentes no arquivo
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            content = file.read()
    except UnicodeDecodeError:
        with open(filepath, 'r', encoding='latin-1') as file:
            content = file.read()
    
    # Padrão para identificar proventos (valores positivos com + no final)
    proventos_pattern = r'(\d{4})\s+([A-Za-zÀ-Ü\s/\(\)\.\-]+)\s+[\d,\.]+\s+[\d,\.]+\s+([\d\.,]+)\+'
    
    # Padrão para identificar descontos (valores negativos com - no final)
    descontos_pattern = r'(\d{4})\s+([A-Za-zÀ-Ü\s/\(\)\.\-]+)\s+[\d,\.]+\s+[\d,\.]+\s+([\d\.,]+)\-'
    
    proventos_encontrados = set()
    descontos_encontrados = set()
    
    # Encontrar proventos
    matches_proventos = re.findall(proventos_pattern, content)
    for codigo, descricao, valor in matches_proventos:
        proventos_encontrados.add((codigo, descricao.strip()))
    
    # Encontrar descontos
    matches_descontos = re.findall(descontos_pattern, content)
    for codigo, descricao, valor in matches_descontos:
        descontos_encontrados.add((codigo, descricao.strip()))
    
    return proventos_encontrados, descontos_encontrados

def extrair_valor_numerico(valor_str):
    """
    Extrai valor numérico de string formatada como 'R$ X.XXX,XX'
    """
    if isinstance(valor_str, str) and 'R$' in valor_str:
        try:
            valor_limpo = valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
            return float(valor_limpo)
        except ValueError:
            return 0.0
    return 0.0

def parse_payroll_file(filepath):
    """
    Função para processar arquivo de folha de pagamento e extrair dados dos funcionários
    """
    global PROVENTOS_CODIGOS, DESCONTOS_CODIGOS
    
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            content = file.read()
    except UnicodeDecodeError:
        with open(filepath, 'r', encoding='latin-1') as file:
            content = file.read()
    
    # Identificar todos os proventos e descontos no arquivo
    todos_proventos, todos_descontos = identificar_proventos_descontos_no_arquivo(filepath)
    PROVENTOS_CODIGOS = todos_proventos
    DESCONTOS_CODIGOS = todos_descontos
    
    # Padrão regex para identificar cada bloco de funcionário
    employee_pattern = r'(\d{4,7})\s+([A-ZÀ-Ü\s\.]+)(?=\s+Admissão:)'
    
    employees = []
    current_employee = {}
    current_matricula = None
    
    # Dividir o conteúdo por linhas
    lines = content.split('\n')
    
    for i, line in enumerate(lines):
        # Procurar início de um novo funcionário
        employee_match = re.search(employee_pattern, line)
        
        if employee_match:
            # Se já temos um funcionário em processamento, adicionar à lista
            if current_employee:
                employees.append(current_employee)
            
            matricula = employee_match.group(1)
            nome = employee_match.group(2).strip()
            
            # Verificar se já existe funcionário com mesma matrícula
            funcionario_existente = next((emp for emp in employees if emp['Matrícula'] == matricula), None)
            
            if funcionario_existente:
                current_employee = funcionario_existente
            else:
                # Inicializar dicionário do funcionário com todos os proventos e descontos
                current_employee = {
                    'Matrícula': matricula,
                    'Nome': nome,
                    'Departamento': '',
                    'Cargo': '',
                    'Salário Base': 'R$ 0,00'
                }
                
                # Inicializar todos os proventos com zero
                for codigo, descricao in todos_proventos:
                    current_employee[descricao] = 'R$ 0,00'
                
                # Inicializar todos os descontos com zero
                for codigo, descricao in todos_descontos:
                    current_employee[descricao] = 'R$ 0,00'
            
            current_matricula = matricula
        
        # Extrair informações básicas do funcionário
        if current_employee:
            if 'Estabelecimento:' in line:
                dep_match = re.search(r'Estabelecimento:\s+\d+\s+([A-Z/]+)', line)
                if dep_match:
                    current_employee['Departamento'] = dep_match.group(1)
            
            if 'Função:' in line:
                cargo_match = re.search(r'Função:\s+([^-]+)', line)
                if cargo_match:
                    current_employee['Cargo'] = cargo_match.group(1).strip()
            
            if 'Salário:' in line:
                salario_match = re.search(r'Salário:\s+([\d\.,]+)', line)
                if salario_match:
                    current_employee['Salário Base'] = f"R$ {salario_match.group(1)}"
            
            # Procurar por proventos (linhas com + no final)
            provento_match = re.search(r'(\d{4})\s+([A-Za-zÀ-Ü\s/\(\)\.\-]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d\.,]+)\+', line)
            if provento_match:
                codigo = provento_match.group(1)
                descricao = provento_match.group(2).strip()
                valor = provento_match.group(5)
                
                # Verificar se esta descrição existe no nosso dicionário
                for prov_codigo, prov_descricao in todos_proventos:
                    if prov_descricao == descricao:
                        # limpar string e converter
                        valor_novo = float(valor.replace(".", "").replace(",", "."))
                        valor_existente = extrair_valor_numerico(current_employee.get(descricao, "R$ 0,00"))
                        soma = valor_existente + valor_novo
                        current_employee[descricao] = f"R$ {soma:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        break
            
            # Procurar por descontos (linhas com - no final)
            desconto_match = re.search(r'(\d{4})\s+([A-Za-zÀ-Ü\s/\(\)\.\-]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d\.,]+)\-', line)
            if desconto_match:
                codigo = desconto_match.group(1)
                descricao = desconto_match.group(2).strip()
                valor = desconto_match.group(5)
                
                # Verificar se esta descrição existe no nosso dicionário
                for desc_codigo, desc_descricao in todos_descontos:
                    if desc_descricao == descricao:
                        valor_novo = float(valor.replace(".", "").replace(",", "."))
                        valor_existente = extrair_valor_numerico(current_employee.get(descricao, "R$ 0,00"))
                        soma = valor_existente + valor_novo
                        current_employee[descricao] = f"R$ {soma:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        break
            
            # Procurar totais
            if 'Tot.Pagamentos:' in line:
                pagamentos_match = re.search(r'Tot\.Pagamentos:\s+([\d\.,]+)', line)
                descontos_match = re.search(r'Tot\.Descontos:\s+([\d\.,]+)', line)
                liquido_match = re.search(r'Líquido:\s+([\d\.,\-]+)', line)
                
                if pagamentos_match:
                    current_employee['Total de Vencimentos'] = f"R$ {pagamentos_match.group(1)}"
                if descontos_match:
                    current_employee['Total de Descontos'] = f"R$ {descontos_match.group(1)}"
                if liquido_match:
                    current_employee['Líquido'] = f"R$ {liquido_match.group(1)}"
    
    # Adicionar o último funcionário
    if current_employee:
        employees.append(current_employee)
    
    # Remover duplicatas baseado na matrícula
    employees_unicos = []
    matriculas_vistas = set()
    
    for emp in employees:
        if emp['Matrícula'] not in matriculas_vistas:
            employees_unicos.append(emp)
            matriculas_vistas.add(emp['Matrícula'])
    
    return employees_unicos

def create_payroll_table(filepath):
    """
    Função principal para criar a tabela de folha de pagamento
    """
    # Processar o arquivo
    employees_data = parse_payroll_file(filepath)
    
    # Criar DataFrame
    df = pd.DataFrame(employees_data)
    
    # Ordenar por nome
    df = df.sort_values('Nome')
    
    return df

def criar_arquivo_excel(payroll_table, caminho_saida):
    """
    Cria arquivo Excel com formatação e múltiplas abas
    """
    try:
        # Criar writer Excel
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            
            # 1. ABA PRINCIPAL - Dados completos com todos os proventos e descontos
            payroll_table.to_excel(writer, sheet_name='Folha Completa', index=False)
            
            # 2. ABA RESUMO POR DEPARTAMENTO
            df_calc = payroll_table.copy()
            
            # Converter todas as colunas monetárias para valores numéricos
            colunas_monetarias = [col for col in df_calc.columns if col not in ['Matrícula', 'Nome', 'Departamento', 'Cargo']]
            for coluna in colunas_monetarias:
                df_calc[f'{coluna}_Num'] = df_calc[coluna].apply(extrair_valor_numerico)
            
            # Agrupar por departamento
            colunas_soma = [f'{col}_Num' for col in colunas_monetarias]
            resumo_departamento = df_calc.groupby('Departamento').agg({
                'Matrícula': 'count',
                **{col: 'sum' for col in colunas_soma}
            }).rename(columns={'Matrícula': 'Qtd Funcionários'}).round(2)
            
            # Renomear colunas para remover _Num
            novo_nome_colunas = {}
            for col in resumo_departamento.columns:
                if col == 'Qtd Funcionários':
                    novo_nome_colunas[col] = col
                else:
                    novo_nome_colunas[col] = col.replace('_Num', '')
            
            resumo_departamento = resumo_departamento.rename(columns=novo_nome_colunas)
            resumo_departamento.to_excel(writer, sheet_name='Resumo por Departamento')
            
            # 3. ABA ESTATÍSTICAS GERAIS
            total_vencimentos = df_calc['Total de Vencimentos_Num'].sum()
            total_descontos = df_calc['Total de Descontos_Num'].sum()
            total_liquido = df_calc['Líquido_Num'].sum()
            
            estatisticas = pd.DataFrame({
                'Metrica': [
                    'Total de Funcionários',
                    'Total de Vencimentos (R$)',
                    'Total de Descontos (R$)',
                    'Total Líquido (R$)',
                    'Média Salarial (R$)',
                    'Maior Salário (R$)',
                    'Menor Salário (R$)',
                    'Média de Descontos (R$)',
                    'Maior Desconto (R$)'
                ],
                'Valor': [
                    len(payroll_table),
                    total_vencimentos,
                    total_descontos,
                    total_liquido,
                    df_calc['Total de Vencimentos_Num'].mean(),
                    df_calc['Total de Vencimentos_Num'].max(),
                    df_calc['Total de Vencimentos_Num'].min(),
                    df_calc['Total de Descontos_Num'].mean(),
                    df_calc['Total de Descontos_Num'].max()
                ]
            })
            
            estatisticas.to_excel(writer, sheet_name='Estatísticas', index=False)
            
            # 4. ABA TOP 10 MAIORES SALÁRIOS
            top10 = payroll_table.copy()
            top10['Valor_Num'] = top10['Total de Vencimentos'].apply(extrair_valor_numerico)
            top10 = top10.nlargest(10, 'Valor_Num')[['Nome', 'Departamento', 'Cargo', 'Total de Vencimentos'] + 
                                                   [col for col in payroll_table.columns if col not in ['Matrícula', 'Nome', 'Departamento', 'Cargo', 'Total de Vencimentos', 'Total de Descontos', 'Líquido']]]
            top10.to_excel(writer, sheet_name='Top 10 Salários', index=False)
            
            # 5. ABA DETALHAMENTO DE PROVENTOS E DESCONTOS
            proventos_descontos_detalhados = []
            for _, row in payroll_table.iterrows():
                for coluna in payroll_table.columns:
                    if coluna not in ['Matrícula', 'Nome', 'Departamento', 'Cargo', 'Total de Vencimentos', 'Total de Descontos', 'Líquido']:
                        valor_numerico = extrair_valor_numerico(row[coluna])
                        if valor_numerico > 0:
                            tipo = 'Provento' if coluna in [desc for cod, desc in PROVENTOS_CODIGOS] else 'Desconto'
                            proventos_descontos_detalhados.append({
                                'Matrícula': row['Matrícula'],
                                'Nome': row['Nome'],
                                'Departamento': row['Departamento'],
                                'Tipo': tipo,
                                'Descrição': coluna,
                                'Valor': row[coluna],
                                'Valor Numérico': valor_numerico
                            })
            
            df_proventos_descontos = pd.DataFrame(proventos_descontos_detalhados)
            df_proventos_descontos.to_excel(writer, sheet_name='Detalhamento', index=False)
            
            # 6. ABA RESUMO DE DESCONTOS
            descontos_resumo = []
            for codigo, descricao in DESCONTOS_CODIGOS:
                if descricao in payroll_table.columns:
                    total_desconto = df_calc[f'{descricao}_Num'].sum()
                    descontos_resumo.append({
                        'Código': codigo,
                        'Descrição': descricao,
                        'Total (R$)': total_desconto,
                        'Quantidade': (df_calc[f'{descricao}_Num'] > 0).sum()
                    })
            
            df_descontos_resumo = pd.DataFrame(descontos_resumo)
            df_descontos_resumo.to_excel(writer, sheet_name='Resumo Descontos', index=False)
            
        print(f"✓ Arquivo Excel criado com sucesso: {caminho_saida}")
        return True
        
    except Exception as e:
        print(f"✗ Erro ao criar arquivo Excel: {e}")
        return False

def salvar_arquivo_csv(payroll_table, caminho_csv):
    """
    Salva os dados em arquivo CSV
    """
    try:
        payroll_table.to_csv(caminho_csv, index=False, encoding='utf-8-sig')
        print(f"✓ Arquivo CSV salvo: {caminho_csv}")
        return True
    except Exception as e:
        print(f"✗ Erro ao salvar CSV: {e}")
        return False

def verificar_dependencias():
    """
    Verifica se todas as dependências estão instaladas
    """
    try:
        import openpyxl
        print("✓ openpyxl está instalado")
        return True
    except ImportError:
        print("✗ openpyxl não está instalado")
        print("Instale com: pip install openpyxl")
        return False

def verificar_arquivo_txt():
    """
    Verifica se o arquivo TXT existe
    """
    if not os.path.exists(CAMINHO_ARQUIVO_TXT):
        print(f"✗ Arquivo não encontrado: {CAMINHO_ARQUIVO_TXT}")
        return False
    print(f"✓ Arquivo encontrado: {CAMINHO_ARQUIVO_TXT}")
    return True

def gerar_relatorio_estatisticas(payroll_table):
    """
    Gera relatório estatístico no console
    """
    print("\n" + "=" * 60)
    print("RELATÓRIO ESTATÍSTICO")
    print("=" * 60)
    
    # Extrair valores numéricos
    df_calc = payroll_table.copy()
    colunas_monetarias = [col for col in df_calc.columns if col not in ['Matrícula', 'Nome', 'Departamento', 'Cargo']]
    for coluna in colunas_monetarias:
        df_calc[f'{coluna}_Num'] = df_calc[coluna].apply(extrair_valor_numerico)
    
    # Estatísticas gerais
    total_vencimentos = df_calc['Total de Vencimentos_Num'].sum()
    total_descontos = df_calc['Total de Descontos_Num'].sum()
    total_liquido = df_calc['Líquido_Num'].sum()
    
    print(f"Total de funcionários: {len(payroll_table)}")
    print(f"Total de vencimentos: R$ {total_vencimentos:,.2f}")
    print(f"Total de descontos: R$ {total_descontos:,.2f}")
    print(f"Total líquido: R$ {total_liquido:,.2f}")
    
    # Estatísticas por departamento
    print("\n" + "-" * 40)
    print("RESUMO POR DEPARTAMENTO:")
    for departamento in df_calc['Departamento'].unique():
        func_dep = df_calc[df_calc['Departamento'] == departamento]
        total_dep = func_dep['Vencimentos_Num'].sum()
        descontos_dep = func_dep['Total de Descontos_Num'].sum()
        print(f"  {departamento}: {len(func_dep)} func - Venc: R$ {total_dep:,.2f} - Desc: R$ {descontos_dep:,.2f}")

def main():
    """
    Função principal de execução do script
    """
    print("INICIANDO PROCESSAMENTO DA FOLHA DE PAGAMENTO")
    print("=" * 60)
    
    # Verificar dependências
    if not verificar_dependencias():
        return
    
    # Verificar arquivo
    if not verificar_arquivo_txt():
        return
    
    try:
        # Processar arquivo
        print("📊 Processando arquivo TXT...")
        payroll_table = create_payroll_table(CAMINHO_ARQUIVO_TXT)
        print(f"✓ Dados processados: {len(payroll_table)} funcionários")
        print(f"✓ Proventos identificados: {len(PROVENTOS_CODIGOS)} tipos")
        print(f"✓ Descontos identificados: {len(DESCONTOS_CODIGOS)} tipos")
        
        # Gerar nome dos arquivos de saída
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_excel = f"folha_pagamento_agosto_2025_{timestamp}.xlsx"
        caminho_excel = os.path.join(CAMINHO_BASE, nome_excel)
        caminho_csv = os.path.join(CAMINHO_BASE, "folha_pagamento_processada.csv")
        
        # Salvar arquivos
        print("💾 Salvando arquivos...")
        salvar_arquivo_csv(payroll_table, caminho_csv)
        criar_arquivo_excel(payroll_table, caminho_excel)
        
        # Gerar relatório
        gerar_relatorio_estatisticas(payroll_table)
        
        print("\n" + "=" * 60)
        print("🎯 PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
        print(f"📁 Arquivos gerados em: {CAMINHO_BASE}")
        print(f"   • Excel: {nome_excel}")
        print(f"   • CSV: folha_pagamento_processada.csv")
        
    except Exception as e:
        print(f"❌ Erro durante o processamento: {e}")
        print(f"Tipo do erro: {type(e).__name__}")

if __name__ == "__main__":
    # Executar função principal
    main()
