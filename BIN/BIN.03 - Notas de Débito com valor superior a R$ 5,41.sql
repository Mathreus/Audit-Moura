WITH DadosBase AS (
    SELECT 
        *,
        CASE 
            WHEN PESO IS NULL OR PESO = 0 THEN NULL -- Realiza um tratamento para evitar divisão por zero
            ELSE CUSTO_FINANCEIRO / PESO  -- Realiza um cálculo do Custo Financeiro / Peso do BIN
        END AS CUSTO_UNITARIO -- O resultado é o custo unitário
    FROM 
        VW_AUDIT_RM_GERENCIAMENTO_ESTOQUE
    WHERE 
        COD_ESTABELECIMENTO = 'R151' -- Altere para filtrar os Distribuidores que serão auditados
        AND DATA BETWEEN '2025-07-01' AND '2025-12-31' -- Selecione o período da auditoria
        AND TIPO_OPERACAO = '1.7.006' -- Operação: Entrada de BIN (Cliente c/financeiro - nota de débito)
)
SELECT 
    *,
    CASE
        WHEN CUSTO_UNITARIO > 5.41 THEN 'Ponto' -- Caso o valor for maior que R$ 5,41, então ponto
        WHEN CUSTO_UNITARIO < 5.41 THEN 'OK' -- Caso o valor for menor que R$ 5,41, então OK
        ELSE 'N/A'
    END AS CLASSIFICACAO
FROM 
    DadosBase
WHERE 
    CUSTO_UNITARIO > 5.41
ORDER BY 
    CUSTO_UNITARIO DESC;
