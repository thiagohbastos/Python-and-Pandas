#%% 180. Consecutive Numbers
from func_pandas import adaptando_modelo

import pandas as pd

def consecutive_numbers(logs: pd.DataFrame) -> pd.DataFrame:
    reference_number = 0
    repeat = 0
    result_list = []
    for tuple in logs.values:
        if reference_number == tuple[1] and repeat == 2:
            result_list.append(tuple[1])
            repeat += 1
        
        elif reference_number == tuple[1] and repeat <2:
            repeat += 1
            continue

        else:
            reference_number = tuple[1]
            repeat = 1
    temp = []
    LC_function = [temp.append(x) for x in result_list if x not in temp]
    result = pd.DataFrame({'ConsecutiveNums': temp})
    return result


#%%
str_num = '''
| id | num |
+----+-----+
| 1  | 1   |
| 2  | 1   |
| 3  | 1   |
| 4  | 2   |
| 5  | 1   |
| 6  | 2   |
| 7  | 2   |
'''
numbers = adaptando_modelo(str_num)


#%%
result = consecutive_numbers(logs=numbers)
