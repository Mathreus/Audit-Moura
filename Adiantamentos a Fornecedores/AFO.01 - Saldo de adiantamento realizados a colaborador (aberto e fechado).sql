-- Data Base
DECLARE @DATA_BASE DATE = '2025-12-31' 

SELECT
    COD_ESTABELECIMENTO,
    COD_FORNECEDOR,
    MAX([CNPJ/CPF]) AS CNPJ_CPF,
    MAX(NOME_FORNECEDOR) AS NOME_FORNECEDOR,
    MAX(PERFIL_LANC) AS PERFIL_LANC,
    -- Soma condicional para o status 'ABERTO'
    SUM(CASE WHEN STATUS = 'ABERTO' THEN VALOR_MOEDA ELSE 0 END) AS VALOR_ABERTO,
    -- Soma condicional para outros status 'FECHADO' 
    SUM(CASE WHEN STATUS = 'FECHADO' THEN VALOR_MOEDA ELSE 0 END) AS VALOR_FECHADO,
    -- Soma valor geral
    SUM(VALOR_MOEDA) AS VALOR_TOTAL,
    -- Quantidade de adiantamento em aberto 
    SUM(CASE WHEN STATUS = 'ABERTO' THEN 1 ELSE 0 END) AS QTD_ABERTO,
    --Quantidade de adiantamento fechados
    SUM(CASE WHEN STATUS = 'FECHADO' THEN 1 ELSE 0 END) AS QTD_FECHADO,
    -- Quantidade Total de adiantamentos
    (SUM(CASE WHEN STATUS = 'ABERTO' THEN 1 ELSE 0 END)) + (SUM(CASE WHEN STATUS = 'FECHADO' THEN 1 ELSE 0 END)) AS QTD_TOTAL
FROM
    VW_AUDIT_RM_TRANSACOES_FORNECEDOR
WHERE
    COD_ESTABELECIMENTO IN ('R351', 'R352') -- Alterar para os Distribuidores que serão auditados
    AND DATA_TRANSACAO <= @DATA_BASE -- Data Base
    AND (DATA_LIQUIDACAO >= DATEADD(DAY, 1, @DATA_BASE) OR DATA_LIQUIDACAO = '1900-01-01') -- Filtro de liquidações que serão liquidadas e ainda não foram liquidadas
    AND PERFIL_LANC = 'AFO' -- Filtro de adiantamento a fornecedores
    AND STATUS IN ('ABERTO', 'FECHADO') -- Filtro de status aberto e fechado
GROUP BY
    COD_ESTABELECIMENTO,
    COD_FORNECEDOR
HAVING
    SUM(VALOR_MOEDA) <> 0 -- Retorna valores diferentes de zero
ORDER BY
    SUM(CASE WHEN STATUS = 'ABERTO' THEN VALOR_MOEDA ELSE 0 END) DESC
