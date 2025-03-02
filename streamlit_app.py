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

    # Initialize Supabase client
    supabase = initialize_supabase()
    if not supabase:
        st.stop()

    # Sidebar for choosing mode
    st.sidebar.title("Mode Selection")
    mode = st.sidebar.radio(
        "Choose Mode",
        ["Upload Excel", "Restore to Supabase"]
    )

    if mode == "Upload Excel":
        upload_excel_mode()
    elif mode == "Restore to Supabase":
        restore_to_database_mode(supabase)

def upload_excel_mode():
    """Mode for uploading Excel files directly."""
    st.header("Upload Excel Files")
    
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
                        "columns": df.columns.tolist(),
                        "data": buffer.getvalue()
                    }
            st.success(f"Loaded {len(uploaded_files)} files with {len(st.session_state.tables)} sheets")

    # Join Configuration and other features remain the same...

def restore_to_database_mode(supabase):
    """Mode for restoring Excel files to a database table with automatic table naming."""
    st.header("Restore Excel to Database")
    
    # Step 1: Upload Excel file
    uploaded_file = st.file_uploader(
        "Choose an Excel file to restore", 
        type="xlsx"
    )
    if uploaded_file is None:
        st.info("Please upload an Excel file to proceed.")
        return

    # Step 2: Parse Excel file
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("Select Sheet", options=sheet_names)
        df = xls.parse(selected_sheet)

        # Preview data
        st.subheader("Preview of Data")
        st.dataframe(df.head())

        # Step 3: Generate table name automatically
        table_name = f"{uploaded_file.name.split('.')[0]}_{selected_sheet}".replace(" ", "_").lower()

        # Step 4: Check if table already exists in Supabase
        response = supabase.table(table_name).select("*").execute()
        if response.get("status_code") == 200 and response.data:
            st.warning(f"Table '{table_name}' already exists in the database. Please rename the file/sheet or clear the existing table.")
            return

        # Step 5: Confirm and restore data
        if st.button("Restore to Database"):
            try:
                # Convert DataFrame to list of dictionaries
                data = df.to_dict(orient="records")
                
                # Insert data into Supabase table
                response = supabase.table(table_name).insert(data).execute()
                if response.get("status_code") == 201:
                    st.success(f"Data successfully restored to table '{table_name}'!")
                else:
                    st.error("Failed to restore data. Please check the table structure and data.")
            except Exception as e:
                st.error(f"Error restoring data: {str(e)}")

    except Exception as e:
        st.error(f"Error parsing Excel file: {str(e)}")

if __name__ == "__main__":
    main()
