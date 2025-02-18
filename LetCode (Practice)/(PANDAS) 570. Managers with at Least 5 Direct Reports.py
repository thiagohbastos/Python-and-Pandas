#%% 570. Managers with at Least 5 Direct Reports

from func_pandas import adaptando_modelo
import numpy as np

import pandas as pd

def find_managers(employee: pd.DataFrame) -> pd.DataFrame:
    if len(employee) == 0:
        return pd.DataFrame({'name':[]})
    managers = pd.DataFrame(employee['managerId'].value_counts()).reset_index()
    managers = managers[managers['count'] >= 5]['managerId'].astype(int)

    result = employee[employee['id'].apply(
        lambda x: x in managers.values)
        ].reset_index(drop=True)

    return result[['name']]



#%%
str_employee = '''
| id  | name  | department | managerId |
| --- | ----- | ---------- | --------- |
| 101 | John  | A          | null      |
| 102 | Dan   | A          | 101       |
| 103 | James | A          | 101       |
| 104 | Amy   | A          | 101       |
| 105 | Anne  | A          | 101       |
| 106 | Ron   | B          | 101       |
| 101 | John  | A          | null      |
| 102 | Dan   | A          | 101       |
| 103 | James | A          | 108       |
| 104 | Amy   | A          | 108       |
| 104 | Amy   | A          | 115       |
| 105 | Anne  | A          | 101       |
| 106 | Ron   | B          | 101       |'''

employee = adaptando_modelo(str_employee)
employee.replace('null', np.NaN, inplace=True)


#%%
result = find_managers(employee)
result
