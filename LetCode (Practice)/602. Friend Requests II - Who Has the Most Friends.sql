/* Write your T-SQL query statement below 
602. Friend Requests II: Who Has the Most Friends*/

WITH TOTAL_ACCEPTED AS
(
	SELECT  requester_id AS ID
	FROM RequestAccepted A

	UNION ALL

	SELECT  accepter_id AS ID
	FROM RequestAccepted A
)

SELECT  TOP (1) ID
       ,COUNT(*) AS NUM
FROM TOTAL_ACCEPTED
GROUP BY ID
ORDER BY 2 DESC
