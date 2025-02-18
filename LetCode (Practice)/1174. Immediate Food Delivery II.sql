--1174. Immediate Food Delivery II
WITH RANK_ORDER AS
(
	SELECT  *
	       ,ROW_NUMBER() OVER (PARTITION BY CUSTOMER_ID ORDER BY ORDER_DATE) AS RANK
	FROM DELIVERY
) 

, RANK_ONE AS
(
	SELECT  *
	FROM RANK_ORDER
	WHERE RANK = 1 
)

SELECT  ROUND(
    SUM(
        CASE WHEN order_date = customer_pref_delivery_date THEN 1.0 
        ELSE 0 END
        ) * 100.0 
    /
    COUNT(*) , 2
    ) AS immediate_percentage
FROM RANK_ONE