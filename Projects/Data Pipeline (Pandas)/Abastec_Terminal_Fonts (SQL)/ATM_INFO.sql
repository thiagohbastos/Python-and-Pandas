SELECT  A.COD_CEN AS TECE
       ,concat(A.NUM_DND,' - ',A.NOME_AGENCIA) AS AGENCIA
       ,A.NUM_DND
       ,A.IDT_TML
FROM [MERCANTIL\B040466].TC_INFO_AG_TSR_TML A