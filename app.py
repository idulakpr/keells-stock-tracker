import streamlit as st
import pandas as pd
import datetime
import os
from supabase import create_client, Client

# --- Streamlit Page Configuration ---
st.set_page_config(
    page_title="Stock History & Analytics System",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Supabase Initialization ---
# Streamlit Secrets (secrets.toml) මගින් Credentials ලබා ගනී
SUPABASE_URL = st.secrets.get("SUPABASE_URL")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("⚠️ Supabase Credentials සොයා ගැනීමට නැත. කරුණාකර `.streamlit/secrets.toml` හෝ Streamlit Cloud Secrets පරීක්ෂා කරන්න.")
    st.stop()

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

# --- Helper Functions ---
def categorize_by_sku(sku):
    """SKU එක අනුව Category එක වෙන් කිරීම"""
    try:
        sku_str = str(sku).strip().split('.')[0]
        if sku_str.isdigit():
            sku_num = int(sku_str)
            if 1000 <= sku_num <= 1999:
                return "Dairies"
            elif 2000 <= sku_num <= 2999:
                return "Rice"
    except Exception:
        pass
    return "Other"

@st.cache_data(ttl=60)
def fetch_all_data():
    """Database එකෙන් සියලුම Data ලබා ගැනීම"""
    try:
        response = supabase.table("stock_history").select("*").execute()
        df = pd.DataFrame(response.data)
        return df
    except Exception as e:
        st.error(f"Error fetching data from Supabase: {e}")
        return pd.DataFrame()

# --- App Header ---
st.title("📦 Integrated Stock Management System")
st.markdown("---")

# --- Tabs Structure ---
tab1, tab2, tab3 = st.tabs([
    "📤 Upload New Stock (Memorize)", 
    "🔍 Outlet Search & History", 
    "🏢 Warehouse Stock (DCW1)"
])

# ==========================================
# TAB 1: UPLOAD & MEMORIZE DATA
# ==========================================
with tab1:
    st.header("Upload Daily Stock Data")
    st.info("මෙහිදී Upload කරන හැම Excel File එකක්ම Database එකේ Timestamp එකක් සමඟ permanently Save වේ.")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        if st.button("🚀 Save & Memorize to Database", type="primary"):
            with st.spinner("Processing and Uploading to Supabase..."):
                try:
                    df = pd.read_excel(uploaded_file)

                    # 🛠️ Fix NaN / Empty values for JSON compatibility (Out of range float fix)
                    df = df.where(pd.notnull(df), None)

                    # SKU Format Cleaning
                    if 'SKU' in df.columns:
                        df['SKU'] = df['SKU'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                    # Category Identification
                    df['Category'] = df['SKU'].apply(categorize_by_sku)

                    # Upload Timestamp
                    upload_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df['Uploaded_At'] = upload_timestamp

                    # Database Column Name Mapping
                    rename_dict = {
                        'SKU Description': 'SKU_Description',
                        'Store Description': 'Store_Description',
                        'Current Stock On Hand Units': 'Current_Stock_Units',
                        'Material Status Description': 'Material_Status_Desc',
                        'Last Update Date Time': 'Last_Update_Time'
                    }
                    df = df.rename(columns=rename_dict)

                    # Keep only required columns that exist in DataFrame
                    db_columns = [
                        'Uploaded_At', 'Plant', 'Store_Description', 'SKU', 
                        'SKU_Description', 'Category', 'Current_Stock_Units', 
                        'Material_Status_Desc', 'Last_Update_Time'
                    ]
                    available_cols = [c for c in db_columns if c in df.columns]
                    df_to_upload = df[available_cols]

                    # Convert DataFrame to Dict/JSON for Supabase Upload
                    records = df_to_upload.to_dict(orient='records')

                    # Upload in batches of 1000 rows (performance optimisation)
                    batch_size = 1000
                    for i in range(0, len(records), batch_size):
                        supabase.table("stock_history").insert(records[i:i+batch_size]).execute()

                    st.balloons()
                    st.success(f"✅ Data successfully Memorized! Upload Batch Timestamp: {upload_timestamp}")
                    st.cache_data.clear()

                except Exception as e:
                    st.error(f"Error Uploading to Database: {e}")

# ==========================================
# FETCH DATA FOR ANALYTICS TABS
# ==========================================
full_df = fetch_all_data()

# ==========================================
# TAB 2: OUTLET SEARCH & ZERO STOCK
# ==========================================
with tab2:
    st.header("Search Outlets & Stock History")

    if full_df.empty:
        st.warning("Database එකේ Data කිසිවක් නැත. කරුණාකර පළමුව Excel File එකක් Upload කරන්න.")
    else:
        # Batch Select Filter
        available_batches = sorted(full_df['Uploaded_At'].dropna().unique(), reverse=True)
        selected_batch = st.selectbox("📅 Select Upload Batch/Timestamp:", available_batches)

        # Filter Data by Batch
        batch_df = full_df[full_df['Uploaded_At'] == selected_batch]

        # Outlet Search Filter
        outlets = sorted(batch_df['Store_Description'].dropna().unique())
        selected_outlet = st.selectbox("🏪 Select Outlet / Store:", ["All Outlets"] + list(outlets))

        filtered_df = batch_df.copy()
        if selected_outlet != "All Outlets":
            filtered_df = filtered_df[filtered_df['Store_Description'] == selected_outlet]

        # Metric Displays
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Records", len(filtered_df))
        col2.metric("Zero Stock SKUs", len(filtered_df[filtered_df['Current_Stock_Units'] == 0]))
        col3.metric("Available Stock SKUs", len(filtered_df[filtered_df['Current_Stock_Units'] > 0]))

        st.subheader("Stock Data View")
        st.dataframe(filtered_df, use_container_width=True)

# ==========================================
# TAB 3: WAREHOUSE STOCK (DCW1)
# ==========================================
with tab3:
    st.header("Warehouse Stock Overview (DCW1 - Kerawalapitiya)")

    if full_df.empty:
        st.warning("Database එකේ Data කිසිවක් නැත.")
    else:
        # Filter for Warehouse plant/store
        dcw1_df = full_df[
            (full_df['Store_Description'].str.contains('DCW1|Kerawalapitiya', case=False, na=False)) |
            (full_df['Plant'].astype(str).str.contains('DCW1', case=False, na=False))
        ]

        if dcw1_df.empty:
            st.info("DCW1 warehouse එකට අදාළ Records හමු වූයේ නැත.")
        else:
            dcw1_batches = sorted(dcw1_df['Uploaded_At'].dropna().unique(), reverse=True)
            latest_dcw1_batch = st.selectbox("📅 Select DCW1 Batch:", dcw1_batches, key="dcw1_batch")
            
            view_dcw1 = dcw1_df[dcw1_df['Uploaded_At'] == latest_dcw1_batch]
            
            st.subheader(f"Warehouse Stock as of {latest_dcw1_batch}")
            st.dataframe(view_dcw1, use_container_width=True)
