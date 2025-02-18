--SQLGDNP, GNU
SELECT  CAST(COD_CEN AS INT)                                                                     AS TECE
       ,CONVERT(DATE,CASE WHEN DATA = 'META' THEN CONVERT(DATE,GETDATE(),105) ELSE NULL END,105) AS DATA
       ,CAST(REPLACE(VALOR,',','.') AS numeric(38,2))                                            AS VLR_TOT
       ,'META DSI'                                                                               AS TRANSACAO
FROM
(
	SELECT  COD_CEN
	       ,META
	FROM [MERCANTIL\B040466].[TI_TEMP_DSI]
	WHERE COD_CEN != 'DATA'
) P2 UNPIVOT (VALOR FOR DATA IN (META)) M
