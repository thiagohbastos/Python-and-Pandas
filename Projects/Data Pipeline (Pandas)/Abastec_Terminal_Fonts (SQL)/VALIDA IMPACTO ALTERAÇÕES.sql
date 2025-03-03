SELECT  A.DATA_REF
       ,SUM(A.AJUSTE_TOTAL)     AS TOTAL_AJUSTE
       ,SUM(A.OTIMIZACAO_TOTAL) AS TOTAL_OTIMIZACAO
       ,SUM(A.VALOR_AJUSTE)     AS TOTAL
	   ,B.TOTAL AS QTD_ATMS_ENVIADOS
	   ,B.ALTERADO AS QTD_ATMS_ALTERADOS
FROM
(
    SELECT  *
           ,CASE WHEN A.META = 0 AND A.VALOR_AJUSTE > 0 THEN VALOR_AJUSTE
                 WHEN A.META > 0 AND A.VALOR_ORIGINAL < 0 AND A.VALOR_TOTAL < 0 AND A.VALOR_AJUSTE > 0 THEN VALOR_AJUSTE
                 WHEN A.META > 0 AND A.VALOR_ORIGINAL < 0 AND A.VALOR_TOTAL > 0 AND A.VALOR_AJUSTE > 0 THEN - A.VALOR_ORIGINAL  ELSE 0 END AS AJUSTE_TOTAL
           ,CASE WHEN A.META > 0 AND A.VALOR_AJUSTE > 0 AND A.VALOR_ORIGINAL < 0 AND A.VALOR_TOTAL > 0 THEN A.VALOR_TOTAL
                 WHEN A.META > 0 AND A.VALOR_AJUSTE > 0 AND A.VALOR_ORIGINAL >= 0 THEN A.VALOR_AJUSTE  ELSE 0 END                          AS OTIMIZACAO_TOTAL
    FROM [MERCANTIL\B042786].TC_HST_ALT_PROG A
    WHERE A.VALOR_AJUSTE > 0 
) A
LEFT JOIN (SELECT 
	CAST(DATA AS date) AS DATA
	,SUM(TOTAL) AS TOTAL
	,SUM(ALTERADO) AS ALTERADO
FROM [MERCANTIL\B042786].TC_VOLUME_ALT_PROG
GROUP BY DATA) B
ON A.DATA_REF = CAST(B.DATA AS DATE)
GROUP BY A.DATA_REF
	,B.TOTAL
	,B.ALTERADO
ORDER BY 1