import streamlit as st
import pandas as pd
from io import BytesIO

def main():
    st.title("Advanced Excel SQL Processor")
    
    # Initialize session state
    if 'tables' not in st.session_state:
        st.session_state.tables = {}
    if 'current_df' not in st.session_state:
        st.session_state.current_df = None
    if 'selected_columns' not in st.session_state:
        st.session_state.selected_columns = []
    if 'filters' not in st.session_state:
        st.session_state.filters = []
    if 'having_clauses' not in st.session_state:
        st.session_state.having_clauses = []

    # Upload Section
    with st.expander("📤 Upload Excel Files", expanded=True):
        uploaded_files = st.file_uploader(
            "Pilih file Excel", 
            type="xlsx",
            accept_multiple_files=True
        )
        if uploaded_files:
            for file in uploaded_files:
                xls = pd.ExcelFile(file)
                for sheet_name in xls.sheet_names:
                    key = f"{file.name} - {sheet_name}"
                    st.session_state.tables[key] = xls.parse(sheet_name)
            st.success(f"Berhasil memuat {len(uploaded_files)} file dengan {len(st.session_state.tables)} sheet")

    # Join Section
    with st.expander("🔗 Join Tables"):
        col1, col2 = st.columns(2)
        with col1:
            left_table = st.selectbox("Tabel Kiri", options=list(st.session_state.tables.keys()))
        with col2:
            right_table = st.selectbox("Tabel Kanan", options=list(st.session_state.tables.keys()))

        left_cols = st.session_state.tables[left_table].columns if left_table else []
        right_cols = st.session_state.tables[right_table].columns if right_table else []

        col1, col2 = st.columns(2)
        with col1:
            left_col = st.selectbox("Kolom Join Tabel Kiri", options=left_cols)
        with col2:
            right_col = st.selectbox("Kolom Join Tabel Kanan", options=right_cols)

        join_type = st.selectbox("Jenis Join", ["inner", "left", "right"])
        
        if st.button("Lakukan Join"):
            try:
                df_left = st.session_state.tables[left_table]
                df_right = st.session_state.tables[right_table]
                
                merged = pd.merge(
                    df_left,
                    df_right,
                    left_on=left_col,
                    right_on=right_col,
                    how=join_type
                )
                
                st.session_state.current_df = merged
                st.session_state.selected_columns = list(merged.columns)
                st.success("Join berhasil!")
                
                st.session_state.filters = []
                st.session_state.having_clauses = []
                
            except Exception as e:
                st.error(f"Error join: {str(e)}")

    # Pemilihan Kolom
    if st.session_state.current_df is not None:
        with st.expander("📋 Pilih Kolom yang Akan Ditampilkan"):
            st.write("**Silakan pilih kolom yang ingin ditampilkan:**")
            all_columns = st.session_state.current_df.columns.tolist()
            
            selected = st.multiselect(
                "Pilih kolom dari tabel hasil join:",
                options=all_columns,
                default=all_columns
            )
            
            if selected:
                st.session_state.selected_columns = selected
                st.session_state.current_df = st.session_state.current_df[selected]
                st.success("Kolom berhasil dipilih!")
            else:
                st.warning("Silakan pilih minimal satu kolom")

    # Filter Section (WHERE)
    if st.session_state.current_df is not None:
        with st.expander("🔍 Filter Data (WHERE Clause)"):
            col1, col2, col3 = st.columns([2,2,4])
            with col1:
                new_filter_col = st.selectbox("Kolom", options=st.session_state.current_df.columns, key="filter_col")
            with col2:
                new_filter_op = st.selectbox("Operator", ["=", ">", "<", ">=", "<=", "<>", "BETWEEN", "LIKE", "IN"], key="filter_op")
            with col3:
                new_filter_val = st.text_input("Nilai", key="filter_val")

            if st.button("Tambahkan Filter"):
                if new_filter_col and new_filter_op and new_filter_val:
                    st.session_state.filters.append((new_filter_col, new_filter_op, new_filter_val))
            
            st.subheader("Filter Aktif")
            for i, (col, op, val) in enumerate(st.session_state.filters):
                st.write(f"{i+1}. {col} {op} {val}")
                if st.button(f"Hapus Filter {i+1}", key=f"remove_filter_{i}"):
                    st.session_state.filters.pop(i)
                    st.experimental_rerun()

            if st.button("Terapkan Semua Filter"):
                try:
                    df = st.session_state.current_df.copy()
                    for col, op, val in st.session_state.filters:
                        if op == "BETWEEN":
                            val1, val2 = map(str.strip, val.split(','))
                            df = df[df[col].between(val1, val2)]
                        elif op == "LIKE":
                            pattern = val.replace("%", ".*").replace("_", ".")
                            df = df[df[col].astype(str).str.contains(pattern)]
                        elif op == "IN":
                            values = list(map(str.strip, val.split(',')))
                            df = df[df[col].isin(values)]
                        else:
                            df = df.query(f"{col} {op} {val}")
                    st.session_state.current_df = df
                    st.success("Filter berhasil diterapkan!")
                except Exception as e:
                    st.error(f"Error filter: {str(e)}")

    # Hasil Akhir
    if st.session_state.current_df is not None:
        st.subheader("Preview Data")
        st.dataframe(st.session_state.current_df.head(100))
        
        # Export
        if st.button("💾 Export ke Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                st.session_state.current_df.to_excel(writer, index=False)
            st.download_button(
                label="Download File Excel",
                data=output.getvalue(),
                file_name="hasil_processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
