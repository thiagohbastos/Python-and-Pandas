#%% 184. Department Highest Salary

import pandas as pd
from func_pandas import adaptando_modelo

def department_highest_salary(employee: pd.DataFrame, department: pd.DataFrame) -> pd.DataFrame:
    unique = employee.merge(department, how='left'
                            , left_on='departmentId'
                            , right_on='id'
                            ,suffixes=('_employee', '_department'))

    max_salary = unique.groupby('departmentId')['salary'].max().reset_index()
    
    result = unique.groupby('departmentId').apply(
        lambda x: x[x['salary'] == x['salary'].max()]
        ).reset_index(drop=True)

    result = result[['name_department', 'name_employee', 'salary']].rename(
        columns= {'name_department': 'Department'
                  , 'name_employee': 'Employee'
                  ,'salary': 'Salary'})
    
    return result



#%%
str_employee = '''
| id | name  | salary | departmentId |
| -- | ----- | ------ | ------------ |
| 1  | Joe   | 70000  | 1            |
| 2  | Jim   | 90000  | 1            |
| 3  | Henry | 80000  | 2            |
| 4  | Sam   | 60000  | 2            |
| 5  | Max   | 90000  | 1            |'''

str_department = '''
| id | name  |
| -- | ----- |
| 1  | IT    |
| 2  | Sales |'''

employee = adaptando_modelo(modelo_pd = str_employee)

department = adaptando_modelo(modelo_pd = str_department)

result = department_highest_salary(employee=employee, department=department)

result.head()
# %%
