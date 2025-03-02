import streamlit as st
import pandas as pd
from io import BytesIO

# Cache data loading to avoid redundant processing
@st.cache_data
def load_table_data(file, sheet_name):
    """Load full table data from an Excel file and return as a DataFrame."""
    return pd.read_excel(file, sheet_name=sheet_name)

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
                    # Store metadata (columns) and compressed data (Parquet format)
                    df = xls.parse(sheet_name)
                    buffer = BytesIO()
                    df.to_parquet(buffer)
                    st.session_state.tables[key] = {
                        "file": file,
                        "sheet_name": sheet_name,
                        "columns": df.columns.tolist(),
                        "data": buffer.getvalue()  # Store data in Parquet format
                    }
            st.success(f"Loaded {len(uploaded_files)} files with {len(st.session_state.tables)} sheets")

    # Join Configuration
    with st.expander("🔗 Configure Join"):
        col1, col2 = st.columns(2)
        with col1:
            left_table = st.selectbox("Left Table", options=list(st.session_state.tables.keys()))
        with col2:
            right_table = st.selectbox("Right Table", options=list(st.session_state.tables.keys()))

        # Get column names from metadata
        left_cols = st.session_state.tables[left_table]["columns"] if left_table in st.session_state.tables else []
        right_cols = st.session_state.tables[right_table]["columns"] if right_table in st.session_state.tables else []

        # Validation messages
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

    # Execute Analysis
    if st.button("🚀 Perform Full Analysis"):
        try:
            progress_bar = st.progress(0)
            
            # Load full table data from Parquet
            def load_table(key):
                buffer = BytesIO(st.session_state.tables[key]["data"])
                return pd.read_parquet(buffer)
            
            progress_bar.progress(20)
            df_left = load_table(left_table)
            df_right = load_table(right_table)

            # Perform join
            merged = pd.merge(
                df_left,
                df_right,
                left_on=left_join_col,
                right_on=right_join_col,
                how=join_type,
                suffixes=('_left', '_right')
            )
            progress_bar.progress(40)

            # Apply column selection
            valid_columns = [col for col in output_columns if col in merged.columns]
            if len(valid_columns) != len(output_columns):
                st.warning("Some selected columns are not available in the merged data. Skipping invalid columns.")
            merged = merged[valid_columns]
            progress_bar.progress(60)

            # Apply filters
            for col, op, val in st.session_state.params['filters']:
                if col in merged.columns:
                    if op == "BETWEEN":
                        try:
                            val1, val2 = map(float, val.split(','))
                            merged = merged[(merged[col] >= val1) & (merged[col] <= val2)]
                        except ValueError:
                            st.error("Invalid numeric values for BETWEEN operation")
                            continue
                    elif op == "LIKE":
                        pattern = val.replace("%", ".*").replace("_", ".")
                        merged = merged[merged[col].astype(str).str.contains(pattern)]
                    elif op == "IN":
                        values = list(map(str.strip, val.split(',')))
                        merged = merged[merged[col].isin(values)]
                    else:
                        try:
                            val = float(val) if merged[col].dtype.kind in 'biufc' else val
                            merged = merged.query(f"`{col}` {op} {val}")
                        except Exception as e:
                            st.error(f"Error applying filter: {str(e)}")
                            continue
            progress_bar.progress(80)

            # Apply aggregation
            group_col = st.session_state.params.get('group_col')
            agg_col = st.session_state.params.get('agg_col')
            agg_func = st.session_state.params.get('agg_func')
            if group_col and agg_col:
                if group_col in merged.columns and agg_col in merged.columns:
                    grouped = merged.groupby(group_col)
                    aggregated = grouped.agg({agg_col: agg_func})
                    
                    # Apply HAVING
                    for col, op, val in st.session_state.params['having_clauses']:
                        if col in aggregated.columns:
                            try:
                                val = float(val) if aggregated[col].dtype.kind in 'biufc' else val
                                aggregated = aggregated.query(f"`{col}` {op} {val}")
                            except Exception as e:
                                st.error(f"Error applying HAVING condition: {str(e)}")
                                continue
                    
                    merged = aggregated.reset_index()
                else:
                    st.warning("Skipping aggregation due to missing columns.")
            progress_bar.progress(100)

            st.session_state.current_df = merged
            st.success("Analysis completed successfully!")
        
        except Exception as e:
            st.error(f"Analysis error: {str(e)}")
            st.stop()

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

    # Reset All
    if st.button("🔄 Reset All"):
        st.session_state.clear()
        st.experimental_rerun()

if __name__ == "__main__":
    main()
