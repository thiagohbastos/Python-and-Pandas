-- 176. Second Highest Salary

WITH highest_salary AS (
    SELECT MAX(salary) AS salary
    FROM Employee
)

SELECT MAX(salary) AS SecondHighestSalary
FROM Employee
WHERE salary < (SELECT salary FROM highest_salary);

/*DECLARE @tuplas INT;
DECLARE @limitador INT = 2;

SELECT @tuplas = COUNT(DISTINCT SALARY)
FROM EMPLOYEE

IF @tuplas < @limitador
BEGIN 
    SELECT NULL AS SecondHighestSalary
END

ELSE 
BEGIN
    WITH RANKING AS (
        SELECT *
            , ROW_NUMBER() OVER(ORDER BY E.SALARY DESC) AS SALARY_RANK
            , COUNT(SALARY) OVER(PARTITION BY SALARY) AS DISTINCTS
        FROM EMPLOYEE E
    )

    SELECT SALARY AS SecondHighestSalary
    FROM RANKING
    WHERE SALARY_RANK = 2
    AND DISTINCTS = 1
END
*/