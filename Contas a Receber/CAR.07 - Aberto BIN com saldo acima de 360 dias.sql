DECLARE @DATA_BASE DATE = '2025-12-31';
SELECT 
    COD_ESTABELECIMENTO,
    DATA_TRANSACAO,
    COD_CLIENTE,
    COMPROVANTE,
    NUM_NOTA_FICAL,
    PESO_BIN,
    LIQUIDACAO_NOTA_FISCAL,
    DATA_VENCIMENTO,
    DATA_LIQUIDACAO,
    NOME_CLIENTE,
    COD_CONGLOMERADO,
    NUMERO_ORIGEM,
    SISTEMA

    -- Transações originais (PESO_BIN positivos até a data base)
    MAX(CASE 
        WHEN DATA_TRANSACAO <= @DATA_BASE 
            AND PESO_BIN > 0  -- Apenas valores positivos (débitos)
        THEN PESO_BIN 
        ELSE 0 
    END) AS PESO_BIN_ORIGINAL,

     -- Baixas (PESO_BIN negativos até a data base)
    ABS(SUM(CASE 
        WHEN DATA_TRANSACAO <= @DATA_BASE 
            AND PESO_BIN < 0  -- Apenas valores negativos (créditos)
        THEN PESO_BIN 
        ELSE 0 
    END)) AS PESO_BIN_BAIXADO

    -- Saldo em aberto (somatório de todos os PESO_BIN até a data base)
    SUM(CASE 
        WHEN DATA_TRANSACAO <= @DATA_BASE 
        THEN PESO_BIN 
        ELSE 0 
    END) AS SALDO_ABERTO
    
FROM 
    VW_AUDIT_RM_TRANSACOES_BIN
WHERE 
    COD_ESTABELECIMENTO = 'R351'
GROUP BY
    COD_ESTABELECIMENTO,
    COD_CLIENTE,
    NOME_CLIENTE,
    NUM_NOTA_FICAL,
    COMPROVANTE
