#%% 577. Employee Bonus
from func_pandas import adaptando_modelo
import pandas as pd

def employee_bonus(employee: pd.DataFrame, bonus: pd.DataFrame) -> pd.DataFrame:
    n_aparecer = bonus[bonus['bonus'] > 1000]
    result = employee.merge(bonus, how='left', on='empId')
    #result = result.fillna(0)
    #result = result[~(result.bonus > 1000)]

    return result


#%%
employee = adaptando_modelo('''
| empId | name   | supervisor | salary |
+-------+--------+------------+--------+
| 3     | Brad   | null       | 4000   |
| 1     | John   | 3          | 1000   |
| 2     | Dan    | 3          | 2000   |
| 4     | Thomas | 3          | 4000   |
''')

bonus = adaptando_modelo('''
| empId | bonus |
+-------+-------+
| 2     | 500   |
| 4     | 2000  |
''')



#%%
result = employee_bonus(employee, bonus)
result.head()
# %%
