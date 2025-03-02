import streamlit as st
import pandas as pd
from io import BytesIO
from supabase import create_client

# Initialize Supabase client using Streamlit Secrets
def initialize_supabase():
    try:
        SUPABASE_URL = st.secrets["supabase"]["url"]
        SUPABASE_KEY = st.secrets["supabase"]["key"]
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
        return supabase
    except Exception as e:
        st.error(f"Failed to initialize Supabase client: {str(e)}")
        return None

# Function to restore data to Supabase
def restore_to_supabase(table_name, dataframe):
    """Restore a DataFrame to a Supabase table."""
    try:
        # Convert DataFrame to list of dictionaries
        data = dataframe.to_dict(orient="records")
        
        # Insert data into Supabase table
        response = supabase.table(table_name).insert(data).execute()
        if response.get("status_code") == 201:
            st.success(f"Data successfully restored to table '{table_name}'!")
        else:
            st.error("Failed to restore data. Please check the table structure and data.")
    except Exception as e:
        st.error(f"Error restoring data: {str(e)}")

def main():
    st.title("Advanced Excel Data Processor with Supabase Integration")
    
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

    # Initialize Supabase client
    supabase = initialize_supabase()
    if not supabase:
        st.stop()

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
                    # Store full data in session state
                    df = xls.parse(sheet_name)
                    st.session_state.tables[key] = {
                        "file_name": file.name,
                        "sheet_name": sheet_name,
                        "data": df
                    }
            st.success(f"Loaded {len(uploaded_files)} files with {len(st.session_state.tables)} sheets")

    # Select Sheet to Restore
    if st.session_state.tables:
        st.subheader("Select Sheet to Restore to Supabase")
        selected_key = st.selectbox("Select Sheet", options=list(st.session_state.tables.keys()))
        selected_data = st.session_state.tables[selected_key]["data"]

        # Preview data
        st.write("Preview of Selected Sheet:")
        st.dataframe(selected_data.head())

        # Generate table name automatically
        table_name = f"{st.session_state.tables[selected_key]['file_name'].split('.')[0]}_{st.session_state.tables[selected_key]['sheet_name']}".replace(" ", "_").lower()

        # Restore button
        if st.button("Restore to Supabase"):
            restore_to_supabase(table_name, selected_data)

    # Join Configuration
    with st.expander("🔗 Configure Join"):
        col1, col2 = st.columns(2)
        with col1:
            left_table = st.selectbox("Left Table", options=list(st.session_state.tables.keys()))
        with col2:
            right_table = st.selectbox("Right Table", options=list(st.session_state.tables.keys()))

        # Ensure left_cols and right_cols are initialized safely
        left_cols = st.session_state.tables[left_table]["data"].columns if left_table in st.session_state.tables else []
        right_cols = st.session_state.tables[right_table]["data"].columns if right_table in st.session_state.tables else []

        # Validation messages
        if not left_table or not right_table:
            st.warning("Please select both Left Table and Right Table to configure the join.")
        elif not left_cols.tolist() or not right_cols.tolist():
            st.warning("Selected tables do not contain any columns. Please upload valid Excel files.")

        col1, col2 = st.columns(2)
        with col1:
            left_join_col = st.selectbox("Left Join Column", options=left_cols)
        with col2:
            right_join_col = st.selectbox("Right Join Column", options=right_cols)

        join_type = st.selectbox("Join Type", ["inner", "left", "right"])
        output_columns = st.multiselect(
            "Select columns to display", 
            options=left_cols.tolist() + right_cols.tolist()
        )

    # Execute Analysis
    if st.button("🚀 Perform Full Analysis"):
        try:
            # Step 1: Join Tables
            df_left = st.session_state.tables[left_table]["data"]
            df_right = st.session_state.tables[right_table]["data"]

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
