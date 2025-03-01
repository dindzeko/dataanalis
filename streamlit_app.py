import streamlit as st
import pandas as pd
from io import BytesIO

def main():
    st.title("Excel SQL-like Processor")
    
    # Initialize session state
    if 'tables' not in st.session_state:
        st.session_state.tables = {}
    if 'current_df' not in st.session_state:
        st.session_state.current_df = None
    
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
                    st.session_state.tables[key] = xls.parse(sheet_name)
            st.success(f"Loaded {len(uploaded_files)} files with {len(st.session_state.tables)} sheets")
    
    # Join Section
    with st.expander("🔗 Join Tables"):
        col1, col2 = st.columns(2)
        with col1:
            left_table = st.selectbox("Left Table", options=list(st.session_state.tables.keys()))
            left_columns = st.session_state.tables[left_table].columns.tolist()
            left_col = st.selectbox("Left Join Column", options=left_columns)
        with col2:
            right_table = st.selectbox("Right Table", options=list(st.session_state.tables.keys()))
            right_columns = st.session_state.tables[right_table].columns.tolist()
            right_col = st.selectbox("Right Join Column", options=right_columns)
        
        join_type = st.selectbox("Join Type", ["inner", "left", "right"])
        if st.button("Perform Join"):
            try:
                df_left = st.session_state.tables[left_table]
                df_right = st.session_state.tables[right_table]
                st.session_state.current_df = pd.merge(
                    df_left,
                    df_right,
                    left_on=left_col,
                    right_on=right_col,
                    how=join_type
                )
                st.success("Join successful!")
            except Exception as e:
                st.error(f"Join error: {str(e)}")
    
    # Filter Section (WHERE clause)
    if st.session_state.current_df is not None:
        with st.expander("FilterWhere"):
            col1, col2, col3 = st.columns([2,2,4])
            with col1:
                filter_col = st.selectbox("Column", options=st.session_state.current_df.columns, key="where_col")
            with col2:
                filter_op = st.selectbox("Operator", ["=", ">", "<", ">=", "<=", "<>", "BETWEEN", "LIKE", "IN"], key="where_op")
            with col3:
                filter_val = st.text_input("Value", key="where_val")
            
            if st.button("Apply WHERE Filter"):
                try:
                    df = st.session_state.current_df
                    if filter_op == "BETWEEN":
                        val1, val2 = map(str.strip, filter_val.split(','))
                        df_filtered = df[df[filter_col].between(val1, val2)]
                    elif filter_op == "LIKE":
                        pattern = filter_val.replace("%", ".*").replace("_", ".")
                        df_filtered = df[df[filter_col].astype(str).str.contains(pattern)]
                    elif filter_op == "IN":
                        values = list(map(str.strip, filter_val.split(',')))
                        df_filtered = df[df[filter_col].isin(values)]
                    else:
                        df_filtered = df.query(f"{filter_col} {filter_op} {filter_val}")
                    
                    st.session_state.current_df = df_filtered
                    st.success("WHERE filter applied!")
                except Exception as e:
                    st.error(f"Filter error: {str(e)}")
    
    # Aggregation Section with HAVING
    if st.session_state.current_df is not None:
        with st.expander("HAVING Filter (Post-Aggregation)"):
            st.subheader("Aggregation Setup")
            group_col = st.selectbox("Group By Column", options=st.session_state.current_df.columns)
            agg_col = st.selectbox("Aggregation Column", options=st.session_state.current_df.columns)
            agg_func = st.selectbox("Function", ["sum", "mean", "count", "min", "max"])
            
            st.subheader("HAVING Clause")
            having_col = st.selectbox("Aggregated Column", options=[agg_col])
            having_op = st.selectbox("Operator", ["=", ">", "<", ">=", "<="], key="having_op")
            having_val = st.text_input("Value", key="having_val")
            
            if st.button("Apply Aggregation & HAVING"):
                try:
                    grouped = st.session_state.current_df.groupby(group_col)
                    aggregated_df = grouped.agg({agg_col: agg_func}).reset_index()
                    
                    if having_col and having_op and having_val:
                        query_str = f"`{having_col}` {having_op} {having_val}"
                        aggregated_df = aggregated_df.query(query_str)
                    
                    st.session_state.current_df = aggregated_df
                    st.success("Aggregation & HAVING applied!")
                except Exception as e:
                    st.error(f"Aggregation error: {str(e)}")
    
    # Sort Section
    if st.session_state.current_df is not None:
        with st.expander("📊 Sort Data"):
            col1, col2 = st.columns(2)
            with col1:
                sort_col = st.selectbox("Sort Column", options=st.session_state.current_df.columns)
            with col2:
                sort_order = st.selectbox("Order", ["Ascending", "Descending"])
            
            if st.button("Apply Sorting"):
                try:
                    st.session_state.current_df = st.session_state.current_df.sort_values(
                        sort_col,
                        ascending=(sort_order == "Ascending")
                    )
                    st.success("Sorting applied!")
                except Exception as e:
                    st.error(f"Sorting error: {str(e)}")
    
    # Results Display
    if st.session_state.current_df is not None:
        st.subheader("Results Preview")
        st.dataframe(st.session_state.current_df.head(100))
        
        # Export
        if st.button("💾 Export to Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                st.session_state.current_df.to_excel(writer, index=False)
            st.download_button(
                label="Download Excel File",
                data=output.getvalue(),
                file_name="processed_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
