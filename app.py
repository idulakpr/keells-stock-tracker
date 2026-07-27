import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime

# Page setup (Mobile friendly)
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🛒 Keells Stock Tracker (with History)")

# --- SUPABASE CONNECTION & ADMIN AUTH SETUP ---
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    
    ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")
except Exception as e:
    st.error("⚠️ Database connection settings (Secrets) සකසා නොමැත!")
    ADMIN_PASSWORD = "admin123"

# --- SESSION STATE FOR ADMIN LOGIN ---
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

def check_admin_password():
    """Admin Access පරික්ෂා කරන Helper Function එක"""
    if st.session_state.is_admin:
        return True
    
    st.info("🔒 මෙම කොටස භාවිත කිරීමට Admin Access අවශ්‍ය වේ.")
    entered_password = st.text_input("Enter Admin Password:", type="password", key="admin_pwd_input")
    
    if st.button("Unlock Admin Features"):
        if entered_password == ADMIN_PASSWORD:
            st.session_state.is_admin = True
            st.success("✅ Admin Authentication Successful!")
            st.rerun()
        else:
            st.error("❌ වැරදි Password එකකි! නැවත උත්සාහ කරන්න.")
    return False

# --- DAIRY SKU CODES ---
DAIRY_SKUS = ['115281', '115282', '115283', '5285', '44132', '126507', '128484', '120115']

def categorize_by_sku(sku):
    if pd.isna(sku):
        return 'Rice'
    sku_clean = str(sku).split('.')[0].strip()
    if sku_clean in DAIRY_SKUS:
        return 'Dairies'
    return 'Rice'

# --- HELPER FUNCTION: SUPABASE PAGINATION FOR DATA FETCHING ---
def fetch_all_batch_data(selected_batch):
    all_rows = []
    page_size = 1000
    start = 0
    
    while True:
        res = supabase.table('stock_history') \
            .select('*') \
            .eq('Uploaded_At', selected_batch) \
            .range(start, start + page_size - 1) \
            .execute()
        
        data = res.data
        if not data:
            break
        all_rows.extend(data)
        if len(data) < page_size:
            break
        start += page_size
        
    return pd.DataFrame(all_rows)

# --- HELPER FUNCTION: FETCH ALL UNIQUE BADGES/TIMESTAMPS ---
def get_unique_badges():
    try:
        response = supabase.table('stock_history').select('Uploaded_At').limit(50000).execute()
        raw_data = response.data
        if raw_data:
            df_temp = pd.DataFrame(raw_data)
            return sorted(df_temp['Uploaded_At'].dropna().unique().tolist(), reverse=True)
        return []
    except Exception as e:
        return []

# --- NAVIGATION MENU ---
st.markdown("### 📌 Navigation Menu")

# Admin Status Banner
if st.session_state.is_admin:
    st.sidebar.success("🔓 Logged in as Admin")
    if st.sidebar.button("🔒 Logout Admin"):
        st.session_state.is_admin = False
        st.rerun()

main_menu = st.radio(
    "ඔයාට අවශ්‍ය Option එක තෝරන්න:",
    [
        "📤 Upload New Stock (Memorize)", 
        "🔍 Outlet Stock Search", 
        "⚠️ Zero Stock Report", 
        "🏬 Warehouse Stock",
        "🗑️ Manage / Delete Uploaded Files"
    ],
    index=1 # Outlets Search එක Default විදිහට තෝරාගෙන ඇත
)

st.markdown("---")

# ================= 1. UPLOAD NEW STOCK (PROTECTED) =================
if main_menu == "📤 Upload New Stock (Memorize)":
    st.subheader("📤 Upload Daily Excel File to Database")
    
    # Passcode Check
    if check_admin_password():
        st.caption("මෙතැනින් ඔයාගේ Stock Date එක සහ Batch එක තෝරලා Excel File එක Upload කරන්න.")

        # --- USER INPUT FOR DATE & BATCH SEQUENCE ---
        col_date, col_slot = st.columns(2)
        with col_date:
            upload_date = st.date_input("📅 Stock Data අදාළ දිනය (Date):", datetime.date.today())
        
        with col_slot:
            batch_num = st.selectbox(
                "🔢 අද දවසේ කීවෙනි Upload එකද? (Batch Number):", 
                ["Batch 1", "Batch 2", "Batch 3", "Batch 4", "Batch 5"]
            )
        
        custom_note = st.text_input("📝 වෙනත් සටහනක් (Optional Note - e.g. Evening Update):", "")

        # Generate Unique Badge Name
        date_str = upload_date.strftime("%Y-%m-%d")
        if custom_note.strip():
            badge_name = f"{date_str} - {batch_num} ({custom_note.strip()})"
        else:
            badge_name = f"{date_str} - {batch_num}"

        st.info(f"🏷️ **මෙම File එක Save වන Badge එක:** `{badge_name}`")

        uploaded_file = st.file_uploader("Choose App.xlsx file", type=["xlsx", "xls"])

        if uploaded_file is not None:
            if st.button("🚀 Save & Memorize to Database"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    status_text.text("📖 Excel File එක කියවමින් පවතී...")
                    df = pd.read_excel(uploaded_file)

                    # 1. Clean Column Names
                    rename_dict = {
                        'SKU Description': 'SKU_Description',
                        'Store Description': 'Store_Description',
                        'Current Stock On Hand Units': 'Current_Stock_Units',
                        'Material Status Description': 'Material_Status_Desc',
                        'Last Update Date Time': 'Last_Update_Time'
                    }
                    df = df.rename(columns=rename_dict)

                    # 2. Clean SKU First before Categorization
                    if 'SKU' in df.columns:
                        df['SKU'] = df['SKU'].astype(str).apply(lambda x: str(x).split('.')[0].strip())

                    # 3. Categorize
                    df['Category'] = df['SKU'].apply(categorize_by_sku)

                    # 4. Stock Column Numeric කිරීම
                    if 'Current_Stock_Units' in df.columns:
                        df['Current_Stock_Units'] = pd.to_numeric(df['Current_Stock_Units'], errors='coerce').fillna(0)

                    # 5. User Badge Name
                    df['Uploaded_At'] = badge_name

                    # 6. Database Columns Matching
                    valid_db_columns = [
                        'Uploaded_At', 'Store', 'Store_Description', 'SKU', 
                        'SKU_Description', 'Category', 'Current_Stock_Units', 
                        'Material_Status_Desc', 'Last_Update_Time'
                    ]
                    
                    cols_to_keep = [c for c in valid_db_columns if c in df.columns]
                    df_upload = df[cols_to_keep].copy()

                    # 7. Clean NaN Values for Supabase JSON Safety
                    df_upload = df_upload.where(pd.notnull(df_upload), None)

                    records = df_upload.to_dict(orient='records')
                    total_records = len(records)
                    
                    status_text.text(f"⬆️ Database එකට Data Upload වෙමින් පවතී... (Total Rows: {total_records})")

                    # Chunk Size 200 to prevent Supabase timeouts
                    chunk_size = 200
                    for i in range(0, total_records, chunk_size):
                        chunk = records[i:i + chunk_size]
                        supabase.table('stock_history').insert(chunk).execute()
                        
                        # Progress bar update
                        progress = min((i + chunk_size) / total_records, 1.0)
                        progress_bar.progress(progress)

                    status_text.empty()
                    progress_bar.empty()
                    
                    st.success(f"✅ Data successfully Memorized under Badge: '{badge_name}'! Total Rows: {total_records}")
                    st.balloons()
                    
                except Exception as e:
                    status_text.empty()
                    progress_bar.empty()
                    st.error(f"❌ Upload එක අසාර්ථක විය! Error Message: {e}")

# ================= 5. MANAGE / DELETE UPLOADED FILES (PROTECTED) =================
elif main_menu == "🗑️ Manage / Delete Uploaded Files":
    st.subheader("🗑️ Upload කරපු Excel Batches අයින් කිරීම")
    
    # Passcode Check
    if check_admin_password():
        timestamps = get_unique_badges()

        if not timestamps:
            st.info("ℹ️ Database එකේ කිසිම Data එකක් නැත.")
        else:
            selected_delete_batch = st.selectbox("❌ Delete කරන්න අවශ්‍ය Upload Batch / Badge එක තෝරන්න:", timestamps)

            st.warning(f"⚠️ ඔබ තෝරාගත් Badge එක (`{selected_delete_batch}`) ස්ථිරවම Database එකෙන් Delete වනු ඇත.")
            
            if st.button("🔴 Delete Selected Batch"):
                with st.spinner("Deleting Batch from Database..."):
                    try:
                        supabase.table('stock_history').delete().eq('Uploaded_At', selected_delete_batch).execute()
                        st.success(f"✅ Badge '{selected_delete_batch}' සාර්ථකව Delete කරන ලදී!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error deleting data: {e}")

# ================= DATA RETRIEVAL LOGIC FOR PUBLIC TABS =================
else:
    timestamps = get_unique_badges()

    if not timestamps:
        st.warning("⚠️ Database එකේ කිසිම Data එකක් නෑ. කරුණාකර පළමුව Excel File එකක් Upload කරන්න.")
    else:
        selected_batch = st.selectbox("📅 Select Stock Upload Batch / Date History:", timestamps)

        with st.spinner("Fetching full stock records from database..."):
            df = fetch_all_batch_data(selected_batch)

        if not df.empty:
            if 'Current_Stock_Units' in df.columns:
                df['Current_Stock_Units'] = pd.to_numeric(df['Current_Stock_Units'], errors='coerce').fillna(0)

            if 'SKU' in df.columns:
                df['SKU'] = df['SKU'].astype(str).apply(lambda x: str(x).split('.')[0].strip())
                df['Category'] = df['SKU'].apply(categorize_by_sku)

            store_desc_col = 'Store_Description' if 'Store_Description' in df.columns else 'Store'
            item_column = 'SKU_Description' if 'SKU_Description' in df.columns else 'SKU'

            warehouse_mask = (
                df[store_desc_col].astype(str).str.contains('DCW1|Kerawalapitiya', case=False, na=False) |
                df.get('Store', pd.Series()).astype(str).str.contains('DCW1', case=False, na=False)
            )
            
            warehouse_df = df[warehouse_mask]
            outlets_df = df[~warehouse_mask]

            # ================= 2. OUTLET SEARCH =================
            if main_menu == "🔍 Outlet Stock Search":
                outlets = sorted(outlets_df[store_desc_col].dropna().unique())
                if outlets:
                    selected_outlet = st.selectbox("📍 Select Outlet / Store", outlets)

                    outlet_data = outlets_df[outlets_df[store_desc_col] == selected_outlet]
                    items = sorted(outlet_data[item_column].dropna().unique())
                    
                    if items:
                        selected_item = st.selectbox("📦 Select Item", items)
                        item_details = outlet_data[outlet_data[item_column] == selected_item].iloc[0]

                        st.markdown("---")
                        st.subheader(f"🔹 {selected_item}")
                        st.info(f"**SKU:** {item_details.get('SKU', 'N/A')}")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"🏢 **Store:** {item_details.get(store_desc_col, 'N/A')}")
                            st.write(f"📊 **Current Stock On Hand:** `{item_details.get('Current_Stock_Units', 0)}` Units")
                            st.write(f"🔄 **System Last Update Time:** {item_details.get('Last_Update_Time', 'N/A')}")

                        with col2:
                            st.write(f"⚙️ **Category:** {item_details.get('Category', 'N/A')}")
                            st.write(f"📝 **Status Description:** {item_details.get('Material_Status_Desc', 'N/A')}")

            # ================= 3. ZERO STOCK REPORT =================
            elif main_menu == "⚠️ Zero Stock Report":
                st.subheader(f"📋 Zero Stock Outlets Report ({selected_batch})")
                sub_tab1, sub_tab2 = st.tabs(["🥛 Dairies", "🍚 Rice"])

                def render_zero_stock_section(category_name):
                    cat_df = outlets_df[outlets_df['Category'] == category_name]
                    zero_df = cat_df[cat_df['Current_Stock_Units'] <= 0]

                    if zero_df.empty:
                        st.success(f"✅ මේ {category_name} Category එකේ කිසිම Outlet එකක් Zero Stock වී නැත.")
                        return

                    cat_items = ["-- All Zero Stock Items --"] + sorted(zero_df[item_column].dropna().unique().tolist())
                    selected_zero_item = st.selectbox(f"🔍 Filter by {category_name} Item (Optional):", cat_items, key=f"zero_{category_name}")

                    display_df = zero_df.copy()
                    
                    if selected_zero_item != "-- All Zero Stock Items --":
                        display_df = display_df[display_df[item_column] == selected_zero_item]
                        outlet_count = len(display_df)
                        st.error(f"🚨 **{selected_zero_item}** Item එක Outlets **{outlet_count}** ක Zero Stock වී ඇත!")
                    else:
                        st.error(f"🚨 Outlets / Items **{len(zero_df)}** ක් Zero Stock වී ඇත!")

                    display_cols = [store_desc_col, 'SKU', item_column, 'Current_Stock_Units', 'Material_Status_Desc']
                    available_disp = [c for c in display_cols if c in display_df.columns]
                    
                    report_df = display_df[available_disp].reset_index(drop=True)
                    report_df.columns = [c.replace('_', ' ') for c in available_disp]

                    st.dataframe(report_df, use_container_width=True)

                with sub_tab1:
                    render_zero_stock_section("Dairies")

                with sub_tab2:
                    render_zero_stock_section("Rice")

            # ================= 4. WAREHOUSE STOCK =================
            elif main_menu == "🏬 Warehouse Stock":
                st.subheader(f"🏬 Warehouse Stock - DCW1 ({selected_batch})")
                
                if not warehouse_df.empty:
                    wh_display_cols = ['SKU', item_column, 'Current_Stock_Units', 'Category']
                    available_wh = [c for c in warehouse_df.columns if c in wh_display_cols]
                    clean_wh_df = warehouse_df[available_wh].reset_index(drop=True)

                    st.dataframe(clean_wh_df, use_container_width=True)
                else:
                    st.warning("⚠️ Warehouse (DCW1) එකට අදාළ Records හමු වූයේ නැත.")
 
