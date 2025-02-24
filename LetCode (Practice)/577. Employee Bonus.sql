/* 577. Employee Bonus */
SELECT E.name,
    B.bonus
FROM EMPLOYEE E LEFT JOIN BONUS B
ON E.empId = B.empId
WHERE (B.bonus < 1000) or (B.bonus is null)