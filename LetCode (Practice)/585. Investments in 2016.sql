-- 585. Investments in 2016
Insurance =
| pid | tiv_2015 | tiv_2016 | lat | lon |
| --- | -------- | -------- | --- | --- |
| 1   | 10       | 5        | 10  | 10  |
| 2   | 20       | 20       | 20  | 20  |
| 3   | 10       | 30       | 20  | 20  |
| 4   | 10       | 40       | 40  | 40  |

-- RESPOSTA MINHA, ANTES DE ESTUDAR O PARTICIONAMENTO DA AGREGAÇÃO EM JANELA
WITH LOCAL_UNI AS (
    SELECT pid
    FROM Insurance
    GROUP BY lat, lon
    HAVING COUNT(*) = 1
)

, INV_REP AS (
    SELECT pid
    FROM Insurance
    GROUP BY tiv_2015
    HAVING COUNT(*) > 1
)

SELECT SUM(tiv_2016) AS tiv_2016
FROM Insurance a
INNER JOIN LOCAL_UNI b ON a.pid = b.pid
INNER JOIN INV_REP c ON A.pid = c.pid


-- RESPOSTA AGREGADA EM JANELA COM OVER

WITH REQUISITOS AS (
    SELECT *
        ,COUNT(*) OVER(PARTITION BY lat, lon) AS UNIQ_LOCAL
        ,COUNT(*) OVER(PARTITION BY tiv_2015) AS DUP_2015
    FROM Insurance
)

SELECT ROUND(SUM(tiv_2016), 2) AS tiv_2016
FROM REQUISITOS
WHERE UNIQ_LOCAL = 1
AND DUP_2015 > 1