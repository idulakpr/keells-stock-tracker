import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime

# Page setup (Mobile friendly)
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🛒 Keells Stock Tracker (with History)")

# --- SUPABASE CONNECTION ---
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    st.error("⚠️ Database connection settings (Secrets) සකසා නොමැත!")

# --- DAIRY SKU CODES (from image) ---
DAIRY_SKUS = ['115281', '115282', '115283', '5285', '44132', '126507', '128484', '120115']

def categorize_by_sku(sku):
    if pd.isna(sku):
        return 'Rice'
    sku_clean = str(sku).split('.')[0].strip()
    if sku_clean in DAIRY_SKUS:
        return 'Dairies'
    return 'Rice'

# --- NAVIGATION MENU ---
st.markdown("### 📌 Navigation Menu")
main_menu = st.radio(
    "ඔයාට අවශ්‍ය Option එක තෝරන්න:",
    [
        "📤 Upload New Stock (Memorize)", 
        "🔍 Outlet Stock Search", 
        "⚠️ Zero Stock Report", 
        "🏬 Warehouse Stock",
        "🗑️ Manage / Delete Uploaded Files"
    ],
    index=0
)

st.markdown("---")

# ================= 1. UPLOAD NEW STOCK =================
if main_menu == "📤 Upload New Stock (Memorize)":
    st.subheader("📤 Upload Daily Excel File to Database")
    st.caption("මෙහිදී Upload කරන හැම Excel එකක්ම Database එකේ Time-stamp එකත් එක්ක Memorize වෙනවා.")

    uploaded_file = st.file_uploader("Choose App.xlsx file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        if st.button("🚀 Save & Memorize to Database"):
            with st.spinner("Processing & Memorizing Data..."):
                try:
                    df = pd.read_excel(uploaded_file)

                    # 1. Clean Column Names First
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

                    # 3. Categorize (Robust Matching)
                    df['Category'] = df['SKU'].apply(categorize_by_sku)

                    # 4. Stock Column එක Numeric කරලා NaN 0 කිරීම
                    if 'Current_Stock_Units' in df.columns:
                        df['Current_Stock_Units'] = pd.to_numeric(df['Current_Stock_Units'], errors='coerce').fillna(0)

                    # 5. Timestamp
                    upload_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df['Uploaded_At'] = upload_timestamp

                    # 6. Database එකේ තියෙන Columns විතරක් Select කරගැනීම
                    valid_db_columns = [
                        'Uploaded_At', 'Store', 'Store_Description', 'SKU', 
                        'SKU_Description', 'Category', 'Current_Stock_Units', 
                        'Material_Status_Desc', 'Last_Update_Time'
                    ]
                    
                    cols_to_keep = [c for c in valid_db_columns if c in df.columns]
                    df_upload = df[cols_to_keep].copy()

                    # 7. JSON NaN Error Fix
                    df_upload['Current_Stock_Units'] = df_upload['Current_Stock_Units'].fillna(0)
                    df_upload = df_upload.astype(object).where(pd.notnull(df_upload), None)

                    records = df_upload.to_dict(orient='records')

                    chunk_size = 500
                    for i in range(0, len(records), chunk_size):
                        chunk = records[i:i + chunk_size]
                        supabase.table('stock_history').insert(chunk).execute()

                    st.success(f"✅ Data successfully Memorized at {upload_timestamp}!")
                    st.balloons()
                except Exception as e:
                    st.error(f"Error Uploading to Database: {e}")

# ================= 5. MANAGE / DELETE UPLOADED FILES =================
elif main_menu == "🗑️ Manage / Delete Uploaded Files":
    st.subheader("🗑️ Upload කරපු Excel Batches අයින් කිරීම")
    try:
        response = supabase.table('stock_history').select('Uploaded_At').execute()
        raw_data = response.data

        if not raw_data:
            st.info("ℹ️ Database එකේ කිසිම Data එකක් නැත.")
        else:
            timestamps = sorted(list(set([r['Uploaded_At'] for r in raw_data if r.get('Uploaded_At')])), reverse=True)
            
            selected_delete_batch = st.selectbox("❌ Delete කරන්න අවශ්‍ය Upload Batch (Timestamp) එක තෝරන්න:", timestamps)

            st.warning(f"⚠️ ඔබ තෝරාගත් Batch එක (`{selected_delete_batch}`) ස්ථිරවම Database එකෙන් Delete වනු ඇත.")
            
            if st.button("🔴 Delete Selected Batch"):
                with st.spinner("Deleting Batch from Database..."):
                    supabase.table('stock_history').delete().eq('Uploaded_At', selected_delete_batch).execute()
                    st.success(f"✅ Batch '{selected_delete_batch}' සාර්ථකව Delete කරන ලදී!")
                    st.rerun()

    except Exception as e:
        st.error(f"Error deleting data: {e}")

# ================= DATA RETRIEVAL LOGIC FOR OTHER TABS =================
else:
    try:
        response = supabase.table('stock_history').select('Uploaded_At').execute()
        raw_data = response.data

        if not raw_data:
            st.warning("⚠️ Database එකේ කිසිම Data එකක් නෑ. කරුණාකර පළමුව Excel File එකක් Upload කරන්න.")
        else:
            timestamps = sorted(list(set([r['Uploaded_At'] for r in raw_data if r.get('Uploaded_At')])), reverse=True)
            selected_batch = st.selectbox("📅 Select Stock Upload Batch/Time History:", timestamps)

            data_resp = supabase.table('stock_history').select('*').eq('Uploaded_At', selected_batch).execute()
            df = pd.DataFrame(data_resp.data)

            if not df.empty:
                # 🛠️ 1. Force Clean Stock Units to Numeric
                if 'Current_Stock_Units' in df.columns:
                    df['Current_Stock_Units'] = pd.to_numeric(df['Current_Stock_Units'], errors='coerce').fillna(0)

                # 🛠️ 2. Force Recalculate Category live on retrieved data using SKU
                if 'SKU' in df.columns:
                    df['Category'] = df['SKU'].astype(str).apply(categorize_by_sku)

                store_desc_col = 'Store_Description' if 'Store_Description' in df.columns else 'Store'
                item_column = 'SKU_Description' if 'SKU_Description' in df.columns else 'SKU'

                # Filter Warehouse Out
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
                                st.write(f"🔄 **Excel Last Update:** {item_details.get('Last_Update_Time', 'N/A')}")

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

                        st.error(f"🚨 Outlets / Items {len(zero_df)} ක් Zero Stock වී ඇත!")

                        cat_items = ["-- All Zero Stock Items --"] + sorted(zero_df[item_column].dropna().unique().tolist())
                        selected_zero_item = st.selectbox(f"🔍 Filter by {category_name} Item (Optional):", cat_items, key=f"zero_{category_name}")

                        display_df = zero_df.copy()
                        if selected_zero_item != "-- All Zero Stock Items --":
                            display_df = display_df[display_df[item_column] == selected_zero_item]

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
                        available_wh = [c for c in wh_display_cols if c in warehouse_df.columns]
                        clean_wh_df = warehouse_df[available_wh].reset_index(drop=True)

                        st.dataframe(clean_wh_df, use_container_width=True)
                    else:
                        st.warning("⚠️ Warehouse (DCW1) එකට අදාළ Records හමු වූයේ නැත.")

    except Exception as e:
        st.error(f"Error fetching data from database: {e}")
