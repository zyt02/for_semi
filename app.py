# app.py

import streamlit as st
import pandas as pd
from io import BytesIO
from analysis import (
    analyze_taxi_expense,
    analyze_schedule_status,
)

st.set_page_config(page_title="Internal Analytics Platform")
st.title("📊 Internal Analytics Tool")

uploaded_file = st.file_uploader("📁 Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file is not None:
    # Read uploaded file
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success("File uploaded successfully!")
        st.subheader("🔍 Data Preview")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # Select analysis type
    analysis_type = st.selectbox("Select analysis type", [
        "Taxi Expense",
        "Schedule Status",
        "Remote Status"
    ])

    # Run analysis
    if st.button("Run Analysis"):
        try:
            if analysis_type == "Taxi Expense":
                excel_data = analyze_taxi_expense(df)
            elif analysis_type == "Schedule Status":
                excel_data = analyze_schedule_status(df)

            st.success("✅ Analysis complete!")

            # excel for download
            st.download_button(
            label=" Download Result as Excel",
            data=excel_data,
            file_name=f"{analysis_type.replace(' ', '_').lower()}_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        except Exception as e:
            st.error(f"❌ Error during analysis: {e}")
