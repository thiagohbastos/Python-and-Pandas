#%% 185. Department Top Three Salaries

from func_pandas import adaptando_modelo

#%%
import pandas as pd

def top_three_salaries(employee: pd.DataFrame, department: pd.DataFrame) -> pd.DataFrame:
    if len(employee) == 0:
        return pd.DataFrame({'Department': []
                             ,'Employee': []
                             , 'Salary': []})

    max_tree_dp = employee[['salary', 'departmentId']].drop_duplicates(
        ).sort_values(by=['departmentId', 'salary'], ascending=[True, False]
        ).reset_index(drop=True)

    max_tree_dp['rank'] = max_tree_dp.groupby('departmentId').cumcount() + 1
    max_tree_dp = max_tree_dp[max_tree_dp['rank'].apply(lambda x: 1 <= x <= 3)]

    employee = employee.merge(department, how='left'
                              , left_on='departmentId'
                              , right_on='id'
                              , suffixes=['', '_dp']
                              )
    
    result = employee.merge(max_tree_dp, how='inner', on=['departmentId', 'salary'])
    result = result[['name_dp', 'name', 'salary']].rename(columns={
        'name_dp': 'Department'
        ,'name': 'Employee'
        ,'salary': 'Salary'
    }).sort_values(by='Department').reset_index(drop=True)

    return result


#%%
str_employee = '''| id | name  | salary | departmentId |
| -- | ----- | ------ | ------------ |
| 1  | Joe   | 85000  | 1            |
| 2  | Henry | 80000  | 2            |
| 3  | Sam   | 60000  | 2            |
| 4  | Max   | 90000  | 1            |
| 5  | Janet | 69000  | 1            |
| 6  | Randy | 85000  | 1            |
| 7  | Will  | 70000  | 1            |'''

str_department = '''| id | name  |
| -- | ----- |
| 1  | IT    |
| 2  | Sales |'''

employee = adaptando_modelo(str_employee)
department = adaptando_modelo(str_department)


#%%
resultado = top_three_salaries(employee, department)
