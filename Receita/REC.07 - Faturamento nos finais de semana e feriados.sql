SELECT 
    COD_CLIENTE,
    COD_ESTABELECIMENTO,
    NOME_CLIENTE,
    [CPF/CNPJ],
    DATA_NOTA_FISCAL,
    NUM_NOTA_FISCAL,
    CFOP,
    SUM(PESO_BIN) AS PESO_BIN_TOTAL,
    SUM(VALOR) AS VALOR_TOTAL,
    DATENAME(WEEKDAY, DATA_NOTA_FISCAL) AS DIA_SEMANA
FROM 
    VW_AUDIT_RM_ORDENS_VENDA
WHERE 
  COD_ESTABELECIMENTO = 'R151'
  AND DATA_NOTA_FISCAL BETWEEN '2024-07-01' AND '2025-07-31'
  AND PARA_FATURAMENTO = 'Sim'
  AND NUM_NOTA_FISCAL NOT LIKE '%EST%'
  AND CFOP IN ('5.101', '5.102', '5.103', '5.104', '5.105', '5.106', '5.107', '5.108', '5.109', 
               '5.110', '5.111', '5.112', '5.113', '5.114', '5.115', '5.116', '5.201', '5.202',
               '5.203', '5.204', '5.205', '5.206', '5.207', '5.208', '5.209', '5.401', '5.402',
               '5.403', '5.404', '5.405', '5.501', '5.502', '5.503', '5.504', '6.101', '6.102',
               '6.103', '6.104', '6.105', '6.106', '6.107', '6.108', '6.109', '6.110', '6.111',
               '6.112', '6.113', '6.114', '6.115', '6.116')
  AND (
    -- Vendas em finais de semana (sábado = 7, domingo = 1)
    DATEPART(WEEKDAY, DATA_NOTA_FISCAL) IN (1, 7)
    OR
    -- Vendas em feriados nacionais
    CAST(DATA_NOTA_FISCAL AS DATE) IN (
        '2025-01-01', '2025-03-03', '2025-03-04', '2025-04-18', 
        '2025-04-21', '2025-05-01', '2025-06-19', '2025-09-07', 
        '2025-10-12', '2025-11-02', '2025-11-20', '2025-12-25'
    )
  )
GROUP BY 
    COD_CLIENTE,
    COD_ESTABELECIMENTO,
    NOME_CLIENTE,
    [CPF/CNPJ],
    DATA_NOTA_FISCAL,
    NUM_NOTA_FISCAL,
    CFOP
ORDER BY 
    COD_CLIENTE,
    DATA_NOTA_FISCAL,
    NUM_NOTA_FISCAL
