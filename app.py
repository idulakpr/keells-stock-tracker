import streamlit as st
import pandas as pd
import datetime
from supabase import create_client, Client

# --- Page Config ---
st.set_page_config(page_title="Stock History System", layout="wide")

# --- Supabase Setup ---
SUPABASE_URL = st.secrets.get("SUPABASE_URL")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("Supabase credentials not found in secrets.toml!")
    st.stop()

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

def categorize_by_sku(sku):
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
    try:
        response = supabase.table("stock_history").select("*").execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Error fetching data: {e}")
        return pd.DataFrame()

st.title("📦 Stock History & Analytics")

tab1, tab2, tab3 = st.tabs(["📤 Upload Stock", "🔍 Outlet Search & Zero Stock", "🏢 Warehouse Stock"])

# --- TAB 1: UPLOAD ---
with tab1:
    st.header("Upload New Excel File")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        if st.button("🚀 Save & Memorize to Database", type="primary"):
            with st.spinner("Uploading data..."):
                try:
                    df = pd.read_excel(uploaded_file)

                    # 🛠️ Float/NaN Out of Range Error එක Fix කරන පේළිය:
                    df = df.where(pd.notnull(df), None)

                    if 'SKU' in df.columns:
                        df['SKU'] = df['SKU'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                    df['Category'] = df['SKU'].apply(categorize_by_sku)

                    upload_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df['Uploaded_At'] = upload_timestamp

                    rename_dict = {
                        'SKU Description': 'SKU_Description',
                        'Store Description': 'Store_Description',
                        'Current Stock On Hand Units': 'Current_Stock_Units',
                        'Material Status Description': 'Material_Status_Desc',
                        'Last Update Date Time': 'Last_Update_Time'
                    }
                    df = df.rename(columns=rename_dict)

                    db_columns = [
                        'Uploaded_At', 'Plant', 'Store_Description', 'SKU', 
                        'SKU_Description', 'Category', 'Current_Stock_Units', 
                        'Material_Status_Desc', 'Last_Update_Time'
                    ]
                    available_cols = [c for c in db_columns if c in df.columns]
                    df_to_upload = df[available_cols]

                    records = df_to_upload.to_dict(orient='records')

                    batch_size = 1000
                    for i in range(0, len(records), batch_size):
                        supabase.table("stock_history").insert(records[i:i+batch_size]).execute()

                    st.balloons()
                    st.success(f"Data saved! Timestamp: {upload_timestamp}")
                    st.cache_data.clear()

                except Exception as e:
                    st.error(f"Error Uploading to Database: {e}")

# --- FETCH DATA ---
full_df = fetch_all_data()

# --- TAB 2: OUTLET SEARCH & ZERO STOCK ---
with tab2:
    st.header("Outlet Search & Zero Stock Report")
    
    if full_df.empty:
        st.warning("No data in Database. Upload an Excel file first.")
    else:
        batches = sorted(full_df['Uploaded_At'].dropna().unique(), reverse=True)
        selected_batch = st.selectbox("📅 Select Upload Batch:", batches)
        
        batch_df = full_df[full_df['Uploaded_At'] == selected_batch]

        col_cat, col_store = st.columns(2)
        
        with col_cat:
            categories = ["All"] + list(batch_df['Category'].dropna().unique())
            selected_cat = st.selectbox("Category Filter:", categories)

        with col_store:
            outlets = ["All"] + list(sorted(batch_df['Store_Description'].dropna().unique()))
            selected_outlet = st.selectbox("Select Outlet / Store:", outlets)

        filtered_df = batch_df.copy()
        if selected_cat != "All":
            filtered_df = filtered_df[filtered_df['Category'] == selected_cat]
        if selected_outlet != "All":
            filtered_df = filtered_df[filtered_df['Store_Description'] == selected_outlet]

        # Zero stock filter option
        show_zero_only = st.checkbox("Show Zero Stock Only (OOS)")
        if show_zero_only:
            filtered_df = filtered_df[filtered_df['Current_Stock_Units'] == 0]

        st.dataframe(filtered_df, use_container_width=True)

# --- TAB 3: WAREHOUSE STOCK (DCW1) ---
with tab3:
    st.header("DCW1 Warehouse Stock")
    if not full_df.empty:
        dcw1_df = full_df[
            (full_df['Store_Description'].str.contains('DCW1|Kerawalapitiya', case=False, na=False)) |
            (full_df['Plant'].astype(str).str.contains('DCW1', case=False, na=False))
        ]
        
        if dcw1_df.empty:
            st.info("No DCW1 records found.")
        else:
            dcw1_batches = sorted(dcw1_df['Uploaded_At'].dropna().unique(), reverse=True)
            selected_dcw1_batch = st.selectbox("Select DCW1 Batch:", dcw1_batches)
            
            view_dcw1 = dcw1_df[dcw1_df['Uploaded_At'] == selected_dcw1_batch]
            st.dataframe(view_dcw1, use_container_width=True)
