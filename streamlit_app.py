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
        with col2:
            right_table = st.selectbox("Right Table", options=list(st.session_state.tables.keys()))

        # Get columns for selected tables
        left_cols = st.session_state.tables[left_table].columns if left_table else []
        right_cols = st.session_state.tables[right_table].columns if right_table else []

        col1, col2 = st.columns(2)
        with col1:
            left_col = st.selectbox("Left Join Column", options=left_cols)
        with col2:
            right_col = st.selectbox("Right Join Column", options=right_cols)

        join_type = st.selectbox("Join Type", ["inner", "left", "right"])
        
        # Column selection for output
        col1, col2 = st.columns(2)
        with col1:
            left_output_cols = st.multiselect("Select Left Table Columns", options=left_cols)
        with col2:
            right_output_cols = st.multiselect("Select Right Table Columns", options=right_cols)

        if st.button("Perform Join"):
            try:
                df_left = st.session_state.tables[left_table]
                df_right = st.session_state.tables[right_table]
                
                # Perform join
                merged = pd.merge(
                    df_left,
                    df_right,
                    left_on=left_col,
                    right_on=right_col,
                    how=join_type
                )
                
                # Select output columns
                output_cols = left_output_cols + right_output_cols
                st.session_state.current_df = merged[output_cols] if output_cols else merged
                st.session_state.selected_columns = list(st.session_state.current_df.columns)
                st.success("Join successful!")
                
                # Reset filters
                st.session_state.filters = []
                st.session_state.having_clauses = []
                
            except Exception as e:
                st.error(f"Join error: {str(e)}")

    # Filter Section (WHERE)
    if st.session_state.current_df is not None:
        with st.expander("🔍 Filter Data (WHERE Clause)"):
            # Add new filter
            col1, col2, col3 = st.columns([2,2,4])
            with col1:
                new_filter_col = st.selectbox("Column", options=st.session_state.current_df.columns, key="filter_col")
            with col2:
                new_filter_op = st.selectbox("Operator", ["=", ">", "<", ">=", "<=", "<>", "BETWEEN", "LIKE", "IN"], key="filter_op")
            with col3:
                new_filter_val = st.text_input("Value", key="filter_val")

            if st.button("Add Filter Condition"):
                if new_filter_col and new_filter_op and new_filter_val:
                    st.session_state.filters.append((new_filter_col, new_filter_op, new_filter_val))
            
            # Show current filters
            st.subheader("Current Filters")
            for i, (col, op, val) in enumerate(st.session_state.filters):
                st.write(f"{i+1}. {col} {op} {val}")
                if st.button(f"Remove Filter {i+1}", key=f"remove_filter_{i}"):
                    st.session_state.filters.pop(i)
                    st.experimental_rerun()

            if st.button("Apply All Filters"):
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
                    st.success("Filters applied!")
                except Exception as e:
                    st.error(f"Filter error: {str(e)}")

    # Aggregation and HAVING Section
    if st.session_state.current_df is not None:
        with st.expander("🧮 Aggregation (GROUP BY & HAVING)"):
            # GROUP BY
            group_col = st.selectbox("Group By Column", options=st.session_state.current_df.columns)
            
            # Aggregation
            col1, col2 = st.columns(2)
            with col1:
                agg_col = st.selectbox("Aggregation Column", options=st.session_state.current_df.columns)
            with col2:
                agg_func = st.selectbox("Function", ["sum", "mean", "count", "min", "max"])
            
            # HAVING Clause
            st.subheader("HAVING Clause")
            col1, col2, col3 = st.columns([2,2,4])
            with col1:
                having_col = st.selectbox("Column", options=st.session_state.current_df.columns, key="having_col")
            with col2:
                having_op = st.selectbox("Operator", ["=", ">", "<", ">=", "<=", "<>"], key="having_op")
            with col3:
                having_val = st.text_input("Value", key="having_val")

            if st.button("Add HAVING Condition"):
                if having_col and having_op and having_val:
                    st.session_state.having_clauses.append((having_col, having_op, having_val))
            
            # Show current HAVING clauses
            st.subheader("Current HAVING Conditions")
            for i, (col, op, val) in enumerate(st.session_state.having_clauses):
                st.write(f"{i+1}. {col} {op} {val}")
                if st.button(f"Remove HAVING {i+1}", key=f"remove_having_{i}"):
                    st.session_state.having_clauses.pop(i)
                    st.experimental_rerun()

            if st.button("Apply Aggregation"):
                try:
                    grouped = st.session_state.current_df.groupby(group_col)
                    aggregated = grouped.agg({agg_col: agg_func}).reset_index()
                    
                    # Apply HAVING clauses
                    for col, op, val in st.session_state.having_clauses:
                        aggregated = aggregated.query(f"{col} {op} {val}")
                    
                    st.session_state.current_df = aggregated
                    st.success("Aggregation applied!")
                except Exception as e:
                    st.error(f"Aggregation error: {str(e)}")

    # Column Selection and Results
    if st.session_state.current_df is not None:
        with st.expander("📊 Final Output Configuration"):
            st.session_state.selected_columns = st.multiselect(
                "Select columns to display and export",
                options=st.session_state.current_df.columns.tolist(),
                default=st.session_state.selected_columns
            )

        # Display Results
        st.subheader("Processed Data Preview")
        if st.session_state.selected_columns:
            st.dataframe(st.session_state.current_df[st.session_state.selected_columns].head(100))
        else:
            st.warning("No columns selected for display")

        # Export
        if st.button("💾 Export to Excel"):
            if st.session_state.selected_columns:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    st.session_state.current_df[st.session_state.selected_columns].to_excel(writer, index=False)
                st.download_button(
                    label="Download Excel File",
                    data=output.getvalue(),
                    file_name="processed_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Please select columns to export")

if __name__ == "__main__":
    main()
