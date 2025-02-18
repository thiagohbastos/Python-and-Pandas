-- 177. Nth Highest Salary
CREATE FUNCTION getNthHighestSalary(@N INT) RETURNS INT AS
BEGIN
    RETURN (
        SELECT MAX(S.SALARY)
        FROM (
            SELECT A.SALARY
                , DENSE_RANK() OVER(ORDER BY A.SALARY DESC) AS RANK
            FROM EMPLOYEE A
        ) S
        WHERE S.RANK = @N
    );
END
