#%% 182. Duplicate Emails
import pandas as pd

def duplicate_emails(person: pd.DataFrame) -> pd.DataFrame:
    
    dp_emails = person[person.duplicated(subset=['email'])][
        ['email']].drop_duplicates()

    return dp_emails
