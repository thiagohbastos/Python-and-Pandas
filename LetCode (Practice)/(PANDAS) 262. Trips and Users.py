#%% 262. Trips and Users
from func_pandas import adaptando_modelo
import pandas as pd

def trips_and_users(trips: pd.DataFrame, users: pd.DataFrame) -> pd.DataFrame:
    if len(trips) == 0:
        return pd.DataFrame({'Day': [], 'Cancellation Rate': []})
    
    banned_ids = users[users['banned'] == 'Yes'][['users_id']]
    trips = trips[
        (trips['request_at'] >= '2013-10-01') &
        (trips['request_at'] <= '2013-10-03') &
        (~trips['client_id'].isin(banned_ids['users_id'])) &
        (~trips['driver_id'].isin(banned_ids['users_id']))
    ]

    result = pd.DataFrame(
        trips.groupby(['request_at', 'status'])['status'].count()
        )
    
    result = result.stack().reset_index()[['request_at', 'status', 0]]
    result['status'] = ['cancelled' if 'cancell' in x else 'compleeted' for x in result['status']]

    total_trips = pd.DataFrame(result.groupby('request_at')[0].sum()).reset_index()

    result = pd.DataFrame(result.groupby(['request_at', 'status'])[0].sum()). \
        stack().reset_index()[['request_at', 'status', 0]]
    result.rename(columns={0:'QTD'}, inplace=True)
    
    result = result[result['status'] == 'cancelled']

    total_trips = total_trips.merge(result, how='left', on='request_at')
    total_trips.rename(columns={'request_at': 'Day', 0:'TOTAL'}, inplace=True)
    
    result = total_trips
    result['Cancellation Rate'] = round(result['QTD'] / result['TOTAL'], 2).fillna(0)
    
    return result[['Day', 'Cancellation Rate']]


#%%
trips = adaptando_modelo('''| id | client_id | driver_id | city_id | status              | request_at |
+----+-----------+-----------+---------+---------------------+------------+
| 1  | 1         | 10        | 1       | completed           | 2013-10-01 |
| 2  | 2         | 11        | 1       | cancelled_by_driver | 2013-10-01 |
| 3  | 3         | 12        | 6       | completed           | 2013-10-01 |
| 4  | 4         | 13        | 6       | cancelled_by_client | 2013-10-01 |
| 5  | 1         | 10        | 1       | completed           | 2013-10-02 |
| 6  | 2         | 11        | 6       | completed           | 2013-10-02 |
| 7  | 3         | 12        | 6       | completed           | 2013-10-02 |
| 8  | 2         | 12        | 12      | completed           | 2013-10-03 |
| 9  | 3         | 10        | 12      | completed           | 2013-10-03 |
| 10 | 4         | 13        | 12      | cancelled_by_driver | 2013-10-03 |''')

users = adaptando_modelo('''
| users_id | banned | role   |
+----------+--------+--------+
| 1        | No     | client |
| 2        | Yes    | client |
| 3        | No     | client |
| 4        | No     | client |
| 10       | No     | driver |
| 11       | No     | driver |
| 12       | No     | driver |
| 13       | No     | driver |
''')


# %%
resultado = trips_and_users(trips, users)
resultado.head(20)


# %%
