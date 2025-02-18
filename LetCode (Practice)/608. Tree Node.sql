/* 608. Tree Node */
WITH CONTADOR AS (
    SELECT DISTINCT p_id,
        COUNT(*) OVER(PARTITION BY p_id) AS QTD
    FROM Tree
)

SELECT id,
    CASE WHEN t.p_id IS NULL THEN 'Root'
    WHEN c.QTD >= 1 THEN 'Inner'
    ELSE 'Leaf' END AS type
FROM Tree t LEFT JOIN CONTADOR c on t.id = c.p_id