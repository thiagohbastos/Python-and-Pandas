/* 1158. Market Analysis I */
WITH ORDERS_BY_ID AS (
    SELECT buyer_id
        , COUNT(*) AS QTD
    FROM ORDERS
    WHERE YEAR(order_date) = 2019
    GROUP BY buyer_id
)

SELECT user_id AS buyer_id
    , join_date
    , CASE WHEN SUM(QTD) IS NULL THEN 0 ELSE SUM(QTD) END AS orders_in_2019
FROM USERS LEFT JOIN ORDERS_BY_ID ON user_id = buyer_id
GROUP BY user_id 
    , join_date
ORDER BY user_id
