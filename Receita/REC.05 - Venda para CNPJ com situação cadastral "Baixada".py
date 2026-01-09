# OBSERVAÇÕES SOBRE O CÓDIGO!

# 1) É necessário baixar a view ORDENS DE VENDA ou o relatório de demontrativo de faturamento
# 2) Realizar uma tabela dinâmica com o código do cliente, nome do cliente e CNPJ
# 3) Copiar e colar a tabela dnâmica em uma outra Aba e deixar como a primeira aba do arquivo em excel
# 4) Salvar o arquivo sem caracteres especiais
# 5) Inserir o caminho onde está salvo o arquivo na linha e inserir /Nome_arquivo.xlsx
# 6) Quando o código começar a rodar, ele vai levar cerca de 20 à 30 segundos para retornar a situação cadastral do CNPJ e trazer, também, a data da situação cadastral
# 7) Deixe o código rodando durante todo dia, as bases geralmente são grande e a velocidade de leitura vai depender da rapidez da internet
# 8) Não execute outro código, pois o python não permite execução de dois códigos ao mesmo tempo, recomendo gerar todos os testes antes de executar este!
# 9) Ao final, Copie e cole todo os resultado em uma planilha de excel, filtre os CNPJ com situação "BAIXADA" e confronte a data da situação com a data de venda
# 10) Caso existam vendas APÓS a data de baixa, se torna ponto de auditoria! Recomendo verificar se o cliente possui títulos em aberto e/ou na PDD.


# Bibliotecas necessárias para execução do código:
import requests
import pandas as pd
import time
import certifi
import urllib3
from datetime import datetime

# Suprime apenas o aviso de SSL inseguro
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Carrega os dados garantindo que CNPJ seja lido como texto
df1 = pd.read_excel(
    r'C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Automações\CNPJ/Cascavel.xlsx',
    dtype={'CNPJ': str}
)

# Remove caracteres não numéricos e preenche com zeros à esquerda
df1['CNPJ'] = df1['CNPJ'].astype(str).str.replace(r'\D', '', regex=True).str.zfill(14)
cnpjs = df1['CNPJ']

# DataFrame para armazenar os resultados
resultados = pd.DataFrame(columns=['CNPJ', 'SITUACAO', 'data_situacao'])

def consulta_cnpj(cnpj):
    cnpj_formatado = str(cnpj).zfill(14)
    url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj_formatado}'
    params = {
        "token": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
        "plugin": "RF"
    }
    
    try:
        # Primeira tentativa com SSL verificado
        response = requests.get(url, params=params, timeout=10, verify=certifi.where())
        response.raise_for_status()
        dados = response.json()
        time.sleep(30)
        
        # Obtém tanto a situação quanto a data_situacao da resposta
        situacao = dados.get('situacao', 'NÃO ENCONTRADO')
        data_situacao = dados.get('data_situacao', 'NÃO DISPONÍVEL')
        return situacao, data_situacao

    except requests.exceptions.SSLError:
        # Segunda tentativa ignorando SSL
        response = requests.get(url, params=params, timeout=10, verify=False)
        response.raise_for_status()
        dados = response.json()
        time.sleep(30)
        
        situacao = dados.get('situacao', 'NÃO ENCONTRADO')
        data_situacao = dados.get('data_situacao', 'NÃO DISPONÍVEL')
        return situacao, data_situacao

    except Exception as e:
        return f"ERRO NA CONSULTA: {str(e)}", "NÃO DISPONÍVEL"

# Processa cada CNPJ
for cnpj in cnpjs:
    situacao, data_situacao = consulta_cnpj(cnpj)
    cnpj_formatado = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
    print(f"{cnpj_formatado} | {situacao} | {data_situacao}")
    resultados.loc[len(resultados)] = [cnpj_formatado, situacao, data_situacao]

# Salva os resultados
resultados.to_excel('Resultados_Situacao_CNPJ.xlsx', index=False)
print("Consulta concluída! Resultados salvos em 'Resultados_Situacao_CNPJ.xlsx'")
