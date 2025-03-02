import streamlit as st
import pandas as pd
import traceback
from io import BytesIO
from pandas.api.types import is_numeric_dtype

# Pastikan sudah install dependencies:
# pip install streamlit pandas openpyxl xlsxwriter

def main():
    st.title("Excel Data Processor Pro")
    
    # Initialize session state
    if 'tables' not in st.session_state:
        st.session_state.tables = {}
    if 'current_df' not in st.session_state:
        st.session_state.current_df = None
    if 'params' not in st.session_state:
        st.session_state.params = {
            'filters': [],
            'having_clauses': [],
            'sort_rules': [],
            'output_columns': [],
            'group_col': None,
            'agg_col': None,
            'agg_func': None
        }

    # Upload Section
    with st.expander("📤 Upload Excel Files", expanded=True):
        uploaded_files = st.file_uploader(
            "Choose Excel files", 
            type="xlsx",
            accept_multiple_files=True,
            key="file_uploader"
        )
        
        if uploaded_files:
            for file in uploaded_files:
                try:
                    xls = pd.ExcelFile(file)
                    for sheet_name in xls.sheet_names:
                        key = f"{file.name} - {sheet_name}"
                        df = xls.parse(sheet_name)
                        # Simpan data dalam format parquet untuk efisiensi
                        buffer = BytesIO()
                        df.to_parquet(buffer)
                        st.session_state.tables[key] = {
                            "columns": df.columns.tolist(),
                            "data": buffer.getvalue()
                        }
                    st.success(f"Loaded: {file.name}")
                except Exception as e:
                    st.error(f"Error reading {file.name}: {str(e)}")

    # Reset All Button
    if st.button("🔄 Reset All"):
        st.session_state.clear()
        st.experimental_rerun()

    # Join Configuration
    with st.expander("🔗 Configure Join"):
        col1, col2 = st.columns(2)
        with col1:
            left_table = st.selectbox(
                "Left Table",
                options=list(st.session_state.tables.keys()),
                key="left_table"
            )
        with col2:
            right_table = st.selectbox(
                "Right Table",
                options=list(st.session_state.tables.keys()),
                key="right_table"
            )

        # Load columns from parquet
        left_cols = []
        right_cols = []
        if left_table in st.session_state.tables:
            left_cols = st.session_state.tables[left_table]["columns"]
        if right_table in st.session_state.tables:
            right_cols = st.session_state.tables[right_table]["columns"]

        col1, col2 = st.columns(2)
        with col1:
            left_join_col = st.selectbox(
                "Left Join Column",
                options=left_cols,
                key="left_join_col"
            )
        with col2:
            right_join_col = st.selectbox(
                "Right Join Column",
                options=right_cols,
                key="right_join_col"
            )

        join_type = st.selectbox(
            "Join Type",
            ["inner", "left", "right"],
            key="join_type"
        )
        
        # Suffix configuration
        col1, col2 = st.columns(2)
        with col1:
            left_suffix = st.text_input("Left Table Suffix", "_left")
        with col2:
            right_suffix = st.text_input("Right Table Suffix", "_right")

    # Column Selection
    with st.expander("📋 Column Selection"):
        merged_cols = left_cols + right_cols
        st.session_state.params['output_columns'] = st.multiselect(
            "Select columns to display",
            options=merged_cols,
            default=merged_cols,
            key="output_columns"
        )

    # Filter Configuration
    with st.expander("🔍 Configure Filters"):
        col1, col2, col3 = st.columns([2,2,4])
        with col1:
            filter_col = st.selectbox(
                "Filter Column",
                options=merged_cols,
                key="filter_col"
            )
        with col2:
            filter_op = st.selectbox(
                "Operator",
                ["=", ">", "<", ">=", "<=", "BETWEEN", "LIKE", "IN"],
                key="filter_op"
            )
        with col3:
            filter_val = st.text_input("Value", key="filter_val")
        
        if st.button("➕ Add Filter"):
            if filter_col and filter_op and filter_val:
                st.session_state.params['filters'].append(
                    (filter_col, filter_op, filter_val)
                )
        
        # Active Filters
        st.subheader("Active Filters")
        for i, (col, op, val) in enumerate(st.session_state.params['filters']):
            cols = st.columns([4,1])
            cols[0].code(f"{col} {op} {val}")
            if cols[1].button("❌", key=f"del_filter_{i}"):
                st.session_state.params['filters'].pop(i)
                st.experimental_rerun()

    # Aggregation Configuration
    with st.expander("🧮 Configure Aggregation"):
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.params['group_col'] = st.selectbox(
                "Group By Column",
                options=merged_cols,
                key="group_col"
            )
        with col2:
            st.session_state.params['agg_col'] = st.selectbox(
                "Aggregation Column",
                options=merged_cols,
                key="agg_col"
            )
        
        st.session_state.params['agg_func'] = st.selectbox(
            "Aggregation Function",
            ["sum", "mean", "count", "min", "max"],
            key="agg_func"
        )
        
        # HAVING Clause
        st.subheader("HAVING Conditions")
        col1, col2, col3 = st.columns([2,2,4])
        with col1:
            having_col = st.selectbox(
                "HAVING Column",
                options=merged_cols,
                key="having_col"
            )
        with col2:
            having_op = st.selectbox(
                "HAVING Operator",
                ["=", ">", "<", ">=", "<="],
                key="having_op"
            )
        with col3:
            having_val = st.text_input("HAVING Value", key="having_val")
        
        if st.button("➕ Add HAVING"):
            if having_col and having_op and having_val:
                st.session_state.params['having_clauses'].append(
                    (having_col, having_op, having_val)
                )
        
        # Active HAVING
        st.subheader("Active HAVING")
        for i, (col, op, val) in enumerate(st.session_state.params['having_clauses']):
            cols = st.columns([4,1])
            cols[0].code(f"{col} {op} {val}")
            if cols[1].button("❌", key=f"del_having_{i}"):
                st.session_state.params['having_clauses'].pop(i)
                st.experimental_rerun()

    # Execute Analysis
    if st.button("🚀 Perform Analysis"):
        try:
            progress_bar = st.progress(0)
            with st.spinner('Processing...'):
                # Load data from parquet
                df_left = pd.read_parquet(
                    BytesIO(st.session_state.tables[left_table]["data"])
                )
                df_right = pd.read_parquet(
                    BytesIO(st.session_state.tables[right_table]["data"])
                )
                progress_bar.progress(20)

                # Perform merge with conflict resolution
                merged = pd.merge(
                    df_left,
                    df_right,
                    left_on=left_join_col,
                    right_on=right_join_col,
                    how=join_type,
                    suffixes=(left_suffix, right_suffix)
                )
                progress_bar.progress(40)

                # Column selection
                if st.session_state.params['output_columns']:
                    valid_cols = [col for col in st.session_state.params['output_columns'] 
                                  if col in merged.columns]
                    merged = merged[valid_cols]
                progress_bar.progress(50)

                # Apply filters
                for col, op, val in st.session_state.params['filters']:
                    if col not in merged.columns:
                        continue
                        
                    if op == "BETWEEN":
                        try:
                            val1, val2 = map(float, val.split(','))
                            merged = merged[(merged[col] >= val1) & 
                                           (merged[col] <= val2)]
                        except:
                            st.error(f"Invalid BETWEEN values: {val}")
                    elif op == "LIKE":
                        pattern = val.replace("%", ".*").replace("_", ".")
                        merged = merged[merged[col].astype(str).str.contains(pattern)]
                    elif op == "IN":
                        values = [x.strip() for x in val.split(',')]
                        merged = merged[merged[col].isin(values)]
                    else:
                        try:
                            if is_numeric_dtype(merged[col]):
                                val = float(val)
                            merged = merged.query(f"`{col}` {op} {val}")
                        except:
                            st.error(f"Invalid filter: {col} {op} {val}")
                progress_bar.progress(60)

                # Apply aggregation
                if st.session_state.params['group_col'] and st.session_state.params['agg_col']:
                    group_col = st.session_state.params['group_col']
                    agg_col = st.session_state.params['agg_col']
                    
                    if group_col in merged.columns:
                        grouped = merged.groupby(group_col)
                        agg_func = st.session_state.params['agg_func']
                        aggregated = grouped.agg({agg_col: agg_func})
                        
                        # Apply HAVING
                        for col, op, val in st.session_state.params['having_clauses']:
                            if col in aggregated.columns:
                                try:
                                    if is_numeric_dtype(aggregated[col]):
                                        val = float(val)
                                    aggregated = aggregated.query(f"`{col}` {op} {val}")
                                except:
                                    st.error(f"Invalid HAVING: {col} {op} {val}")
                        merged = aggregated.reset_index()
                progress_bar.progress(80)

                # Apply sorting
                if st.session_state.params['sort_rules']:
                    sort_cols = []
                    sort_asc = []
                    for col, asc in st.session_state.params['sort_rules']:
                        if col in merged.columns:
                            sort_cols.append(col)
                            sort_asc.append(asc)
                    if sort_cols:
                        merged = merged.sort_values(sort_cols, ascending=sort_asc)
                
                st.session_state.current_df = merged
                progress_bar.progress(100)
                st.success("Analysis completed!")

        except Exception as e:
            st.error(f"""
            **Analysis Failed!**
            Error Details:
            ```
            {traceback.format_exc()}
            ```
            """)
            st.stop()

    # Display Results
    if st.session_state.current_df is not None:
        st.subheader("Results")
        st.dataframe(st.session_state.current_df.head(100))
        
        # Export
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
                st.error(f"Export failed: {str(e)}")

if __name__ == "__main__":
    main()
