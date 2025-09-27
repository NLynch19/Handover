import pandas as pd

def load_excel(path, columns):
    try:
        return pd.read_excel(path)
    except FileNotFoundError:
        return pd.DataFrame(columns=columns)

def save_to_excel(df, path):
    df.to_excel(path, index=False)

