/* Write your T-SQL query statement below 
550. Game Play Analysis IV
*/

-- QUERY QUE TRAZ O PRIMEIRO LOGGIN POR ID
WITH FIRST_LOGGED AS
(
    SELECT  
        player_id,
        MIN(event_date) AS first
    FROM ACTIVITY
    GROUP BY player_id
)

-- QUERY QUE TRAZ A QUANTIDADE DE PLAYERS UNICOS QUE LOGARAM NO DIA SEGUINTE
, GROUPED AS
(
    SELECT 
        COUNT(DISTINCT FIRST_LOGGED.player_id) AS QTD
    FROM FIRST_LOGGED
    INNER JOIN ACTIVITY ON FIRST_LOGGED.player_id = ACTIVITY.player_id
    WHERE DATEDIFF(DAY, FIRST_LOGGED.first, ACTIVITY.event_date) = 1
)

-- QUERY QUE TRAZ O TOTAL DE PLAYERS ÚNICOS
, TOTAL_PLAYERS AS 
(
    SELECT 
        COUNT(DISTINCT player_id) AS total_players
    FROM ACTIVITY
)

SELECT 
    ROUND((GROUPED.QTD * 1.0) / TOTAL_PLAYERS.total_players, 2) AS fraction
FROM GROUPED, TOTAL_PLAYERS;
