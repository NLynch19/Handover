import pandas as pd
import os
import streamlit as st

def load_excel(path, columns):
    """
    Loads an Excel file from the given path.
    If the file is missing or invalid, returns an empty DataFrame with the expected columns.
    """
    try:
        df = pd.read_excel(path)
        st.write(f"📋 {len(df.columns)} columns detected:", df.columns.tolist())
        return df
    except FileNotFoundError:
        st.error(f"⚠️ File not found: {path}. Creating empty DataFrame.")
        return pd.DataFrame(columns=columns)
    except Exception as e:
        st.error(f"❌ Error loading Excel file: {e}")
        return pd.DataFrame(columns=columns)

def save_excel(df, path):
    """
    Saves the given DataFrame to an Excel file at the specified path.
    """
    try:
        df.to_excel(path, index=False)
        st.success(f"✅ Data saved to {path}")
    except Exception as e:
        st.error(f"❌ Failed to save Excel file: {e}")
