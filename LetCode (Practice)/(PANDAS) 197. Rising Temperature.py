#%% 197. Rising Temperature
import pandas as pd

def rising_temperature(weather: pd.DataFrame) -> pd.DataFrame:
    temp = weather.copy()

    temp.sort_values(by='recordDate', inplace=True)
    temp['datediff'] = (temp['recordDate'] - temp['recordDate'].shift(1)).dt.days
    temp['yesterday_value'] = temp['temperature'].shift(1)
    temp = temp.query('yesterday_value < temperature and datediff == 1')[['id']]
    return temp
