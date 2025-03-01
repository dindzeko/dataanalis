import streamlit as st
import pandas as pd

# Upload File Excel
uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx", "xls"], accept_multiple_files=True)
dataframes = {}

if uploaded_files:
    for file in uploaded_files:
        try:
            excel = pd.ExcelFile(file)
            for sheet in excel.sheet_names:
                df = excel.parse(sheet)
                dataframes[f"{file.name}_{sheet}"] = df
        except Exception as e:
            st.error(f"Error reading {file.name}: {str(e)}")

    # Tampilkan daftar tabel
    st.write("Available Tables:")
    selected_table = st.selectbox("Select Table", list(dataframes.keys()))
    st.dataframe(dataframes[selected_table])

# Contoh fitur JOIN (sederhana)
if len(dataframes) >= 2:
    st.header("Join Tables")
    table1 = st.selectbox("Table 1", list(dataframes.keys()))
    table2 = st.selectbox("Table 2", list(dataframes.keys()))
    join_type = st.selectbox("Join Type", ["INNER", "LEFT", "RIGHT"])
    
    if st.button("Perform Join"):
        try:
            merged = pd.merge(
                dataframes[table1],
                dataframes[table2],
                how=join_type.lower()
            )
            st.dataframe(merged)
            dataframes["merged_result"] = merged  # Simpan hasil join
        except Exception as e:
            st.error(f"Join failed: {str(e)}")

# Export hasil
if "merged_result" in dataframes:
    st.download_button(
        "Export to Excel",
        data=dataframes["merged_result"].to_csv(index=False),
        file_name="result.csv"
    )
