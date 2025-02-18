#%% 183. Customers Who Never Order

from func_pandas import adaptando_modelo
import pandas as pd

def find_customers(customers: pd.DataFrame, orders: pd.DataFrame) -> pd.DataFrame:
    if len(customers) == 0:
        return pd.DataFrame({'Customers': []})
    #result = customers.merge(orders, how='left', left_on='id', right_on='customerId')
    #result = result[result['id_y'].isnull()][['name']].rename(columns={'name': 'Customers'})
    result = customers[
        customers['id'].apply(lambda x: x not in orders['customerId'].values)
        ][['name']].rename(columns={'name': 'Customers'})
    
    return result



# %%
customers = adaptando_modelo('''| id | name  |
+----+-------+
| 1  | Joe   |
| 2  | Henry |
| 3  | Sam   |
| 4  | Max   |''')

orders = adaptando_modelo('''| id | customerId |
+----+------------+
| 1  | 3          |
| 2  | 1          |''')


#%%
resultado = find_customers(customers=customers, orders=orders)
resultado.head()