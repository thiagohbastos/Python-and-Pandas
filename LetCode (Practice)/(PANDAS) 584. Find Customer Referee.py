#%% 584. Find Customer Referee
from func_pandas import adaptando_modelo
import pandas as pd

def find_customer_referee(customer: pd.DataFrame) -> pd.DataFrame:
    remove = customer[customer['referee_id'] == 2 ][['id']]
    result = customer[customer['id'].apply(lambda x: x not in remove['id'].values)]
    return result[['name']]


#%% Modelos
customer = adaptando_modelo('''| id | name | referee_id |
+----+------+------------+
| 1  | Will | null       |
| 2  | Jane | null       |
| 3  | Alex | 2          |
| 4  | Bill | null       |
| 5  | Zack | 1          |
| 6  | Mark | 2          |''')


#%%
resultado = find_customer_referee(customer)
resultado.head()