--SQLGDNP, GNU
SELECT  COD_CEN
       ,BANCO
       ,DES_JST
       ,TRANSACAO
       ,OS
       ,SUM(VALOR_TOTAL) AS VALOR
       ,ATUALIZACAO
       ,DTA_PRG
FROM [MERCANTIL\B039918].TI_OS_MOV_NUM
WHERE (TRANSACAO = 'Banco do Brasil' AND DTA_PRG = CAST(GETDATE() AS DATE) ) OR TRANSACAO != 'Banco do Brasil'
GROUP BY  COD_CEN
         ,BANCO
         ,DES_JST
         ,TRANSACAO
         ,OS
         ,ATUALIZACAO
         ,DTA_PRG
ORDER BY COD_CEN
