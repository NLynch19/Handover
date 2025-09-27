import streamlit as st

def show_unsaved_prompt():
    st.warning("⚠️ You have unsaved changes.")
    return st.button("💾 Save to Excel")

