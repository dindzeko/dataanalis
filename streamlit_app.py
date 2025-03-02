import streamlit as st
import pandas as pd
from io import BytesIO

# Pastikan sudah install dependencies:
# pip install streamlit pandas openpyxl xlsxwriter

def main():
    st.title("Advanced Excel Data Processor")
    
    # Initialize session state
    if 'tables' not in st.session_state:
        st.session_state.tables = {}
    if 'current_df' not in st.session_state:
        st.session_state.current_df = None
    if 'params' not in st.session_state:
        st.session_state.params = {
            'selected_columns': [],
            'filters': [],
            'having_clauses': [],
            'sort_rules': []
        }

    # Upload Section
    with st.expander("📤 Upload Excel Files", expanded=True):
        uploaded_files = st.file_uploader(
            "Choose Excel files", 
            type="xlsx",
            accept_multiple_files=True
        )
        if uploaded_files:
            for file in uploaded_files:
                xls = pd.ExcelFile(file)
                for sheet_name in xls.sheet_names:
                    key = f"{file.name} - {sheet_name}"
                    # Simpan metadata (nama kolom) saja, bukan data penuh
                    st.session_state.tables[key] = {
                        "file": file,
                        "sheet_name": sheet_name,
                        "columns": xls.parse(sheet_name).columns.tolist()
                    }
            st.success(f"Loaded {len(uploaded_files)} files with {len(st.session_state.tables)} sheets")

    # Join Configuration
    with st.expander("🔗 Configure Join"):
        col1, col2 = st.columns(2)
        with col1:
            left_table = st.selectbox("Left Table", options=list(st.session_state.tables.keys()))
        with col2:
            right_table = st.selectbox("Right Table", options=list(st.session_state.tables.keys()))

        # Ambil nama kolom dari metadata
        left_cols = st.session_state.tables[left_table]["columns"] if left_table in st.session_state.tables else []
        right_cols = st.session_state.tables[right_table]["columns"] if right_table in st.session_state.tables else []

        # Validasi tabel dan kolom
        if not left_table or not right_table:
            st.warning("Please select both Left Table and Right Table to configure the join.")
        elif not left_cols or not right_cols:
            st.warning("Selected tables do not contain any columns. Please upload valid Excel files.")

        col1, col2 = st.columns(2)
        with col1:
            left_join_col = st.selectbox("Left Join Column", options=left_cols)
        with col2:
            right_join_col = st.selectbox("Right Join Column", options=right_cols)

        join_type = st.selectbox("Join Type", ["inner", "left", "right"])
        output_columns = st.multiselect(
            "Select columns to display", 
            options=left_cols + right_cols
        )

    # Filter Configuration (WHERE Clause)
    with st.expander("🔍 Configure Filters (WHERE)"):
        col1, col2, col3 = st.columns([2, 2, 4])
        with col1:
            filter_col = st.selectbox("Filter Column", options=left_cols + right_cols)
        with col2:
            filter_op = st.selectbox("Operator", ["=", ">", "<", ">=", "<=", "BETWEEN", "LIKE", "IN"])
        with col3:
            filter_val = st.text_input("Value")
        
        if st.button("Add Filter"):
            if filter_col and filter_op and filter_val:
                st.session_state.params['filters'].append((filter_col, filter_op, filter_val))
        
        st.subheader("Active Filters")
        for i, (col, op, val) in enumerate(st.session_state.params['filters']):
            st.write(f"{i+1}. {col} {op} {val}")
            if st.button(f"Remove Filter {i+1}", key=f"remove_filter_{i}"):
                st.session_state.params['filters'].pop(i)
                st.experimental_rerun()

    # Aggregation Configuration
    with st.expander("🧮 Configure Aggregation (GROUP BY & HAVING)"):
        col1, col2 = st.columns(2)
        with col1:
            group_col = st.selectbox("Group By Column", options=left_cols + right_cols)
        with col2:
            agg_col = st.selectbox("Aggregation Column", options=left_cols + right_cols)
        
        agg_func = st.selectbox("Aggregation Function", ["sum", "mean", "count", "min", "max"])
        
        st.subheader("HAVING Clause")
        col1, col2, col3 = st.columns([2, 2, 4])
        with col1:
            having_col = st.selectbox("HAVING Column", options=left_cols + right_cols)
        with col2:
            having_op = st.selectbox("HAVING Operator", ["=", ">", "<", ">=", "<="])
        with col3:
            having_val = st.text_input("HAVING Value")
        
        if st.button("Add HAVING Condition"):
            if having_col and having_op and having_val:
                st.session_state.params['having_clauses'].append((having_col, having_op, having_val))
        
        st.subheader("Active HAVING Conditions")
        for i, (col, op, val) in enumerate(st.session_state.params['having_clauses']):
            st.write(f"{i+1}. {col} {op} {val}")
            if st.button(f"Remove HAVING {i+1}", key=f"remove_having_{i}"):
                st.session_state.params['having_clauses'].pop(i)
                st.experimental_rerun()

    # Sorting Configuration
    with st.expander("📊 Configure Sorting"):
        col1, col2 = st.columns(2)
        with col1:
            sort_col = st.selectbox("Sort Column", options=left_cols + right_cols)
        with col2:
            sort_order = st.selectbox("Sort Order", ["Ascending", "Descending"])
        
        if st.button("Add Sort Rule"):
            st.session_state.params['sort_rules'].append((sort_col, sort_order == "Ascending"))
        
        st.subheader("Active Sort Rules")
        for i, (col, asc) in enumerate(st.session_state.params['sort_rules']):
            st.write(f"{i+1}. {col} {'Ascending' if asc else 'Descending'}")
            if st.button(f"Remove Sort {i+1}", key=f"remove_sort_{i}"):
                st.session_state.params['sort_rules'].pop(i)
                st.experimental_rerun()

    # Execute Analysis
    if st.button("🚀 Perform Full Analysis"):
        try:
            # Muat data tabel lengkap hanya saat analisis dimulai
            df_left = pd.read_excel(st.session_state.tables[left_table]["file"], sheet_name=st.session_state.tables[left_table]["sheet_name"])
            df_right = pd.read_excel(st.session_state.tables[right_table]["file"], sheet_name=st.session_state.tables[right_table]["sheet_name"])

            # Handle overlapping column names by suffixing them
            merged = pd.merge(
                df_left,
                df_right,
                left_on=left_join_col,
                right_on=right_join_col,
                how=join_type,
                suffixes=('_left', '_right')  # Add suffixes to avoid column name conflicts
            )
            
            # Debug: Show the structure of the merged DataFrame
            st.write("Merged DataFrame Preview:")
            st.write(merged.head())

            # Step 2: Apply Column Selection
            if output_columns:
                # Ensure selected columns exist in the merged DataFrame
                valid_columns = [col for col in output_columns if col in merged.columns]
                if len(valid_columns) != len(output_columns):
                    st.warning("Some selected columns are not available in the merged data. Skipping invalid columns.")
                merged = merged[valid_columns]
            
            # Step 3: Apply Filters (WHERE Clause)
            for col, op, val in st.session_state.params['filters']:
                if col not in merged.columns:
                    st.warning(f"Filter column '{col}' not found in the merged data. Skipping this filter.")
                    continue

                if op == "BETWEEN":
                    val1, val2 = map(str.strip, val.split(','))
                    merged = merged[merged[col].between(float(val1), float(val2))] 
                elif op == "LIKE":
                    pattern = val.replace("%", ".*").replace("_", ".")
                    merged = merged[merged[col].astype(str).str.contains(pattern)]
                elif op == "IN":
                    values = list(map(str.strip, val.split(',')))
                    merged = merged[merged[col].isin(values)]
                else:
                    merged = merged.query(f"`{col}` {op} {val}")  # Use backticks to handle special characters in column names
            
            # Step 4: Apply Aggregation
            group_col = st.session_state.params.get('group_col')
            agg_col = st.session_state.params.get('agg_col')
            agg_func = st.session_state.params.get('agg_func')

            if group_col and agg_col:
                # Check if group_col exists in the merged DataFrame
                if group_col not in merged.columns: 
                    # Try to find the correct suffixed column
                    possible_group_cols = [col for col in merged.columns if group_col in col]
                    if possible_group_cols:
                        group_col = possible_group_cols[0]  # Use the first match
                        st.info(f"Using suffixed column '{group_col}' for Group By.")
                    else:
                        st.warning(f"Group By column '{group_col}' not found in the merged data. Skipping aggregation.")
                        group_col = None

                if group_col and agg_col in merged.columns:
                    grouped = merged.groupby(group_col)
                    aggregated = grouped.agg({agg_col: agg_func})
                    
                    # Apply HAVING
                    for col, op, val in st.session_state.params['having_clauses']:
                        if col not in aggregated.columns:
                            st.warning(f"HAVING column '{col}' not found in the aggregated data. Skipping this condition.")
                            continue
                        aggregated = aggregated.query(f"`{col}` {op} {val}")  # Use backticks for safety
                    
                    merged = aggregated.reset_index()
                else:
                    st.warning("Skipping aggregation due to missing columns.")
            
            # Step 5: Apply Sorting
            if st.session_state.params['sort_rules']:
                sort_cols = [col for col, _ in st.session_state.params['sort_rules'] if col in merged.columns]
                sort_asc = [asc for _, asc in st.session_state.params['sort_rules'] if col in merged.columns]
                
                if not sort_cols:
                    st.warning("No valid sort columns found in the merged data. Skipping sorting.")
                else:
                    merged = merged.sort_values(by=sort_cols, ascending=sort_asc)
            
            st.session_state.current_df = merged
            st.success("Analysis completed successfully!")
        
        except Exception as e:
            st.error(f"Analysis error: {str(e)}")

    # Display Results
    if st.session_state.current_df is not None:
        st.subheader("Analysis Results")
        st.dataframe(st.session_state.current_df.head(100)) 
        
        # Export to Excel
        if st.button("💾 Export to Excel"):
            try:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    st.session_state.current_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="Download Excel File",
                    data=output.getvalue(),
                    file_name="analysis_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Export error: {str(e)}")

if __name__ == "__main__":
    main()
