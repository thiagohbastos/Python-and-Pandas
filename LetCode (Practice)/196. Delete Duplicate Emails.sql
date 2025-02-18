/* Write your T-SQL query statement below 
196. Delete Duplicate Emails */

WITH rank AS (
    SELECT *
        ,ROW_NUMBER() OVER(PARTITION BY P.EMAIL ORDER BY ID) rank
    FROM PERSON P
)

DELETE P
FROM PERSON P JOIN rank r ON P.id = r.id and P.email = r.email
WHERE r.rank >= 2