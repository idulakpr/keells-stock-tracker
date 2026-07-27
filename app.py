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

# --- DAIRY SKU CODES ---
DAIRY_SKUS = ['115281', '115282', '115283', '5285', '44132', '126507', '128484', '120115']

def categorize_by_sku(sku):
    sku_val = str(sku).replace('.0', '').strip()
    if sku_val in DAIRY_SKUS:
        return 'Dairies'
    return 'Rice'

# --- NAVIGATION MENU ---
st.markdown("### 📌 Navigation Menu")
main_menu = st.radio(
    "ඔයාට අවශ්‍ය Option එක තෝරන්න:",
    ["📤 Upload New Stock (Memorize)", "🔍 Outlet Stock Search", "⚠️ Zero Stock Report", "🏬 Warehouse Stock"],
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

                    # 🛠️ Float/NaN Out of range fix
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

                    records = df.to_dict(orient='records')

                    chunk_size = 500
                    for i in range(0, len(records), chunk_size):
                        chunk = records[i:i + chunk_size]
                        supabase.table('stock_history').insert(chunk).execute()

                    st.success(f"✅ Data successfully Memorized at {upload_timestamp}!")
                    st.balloons()
                except Exception as e:
                    st.error(f"Error Uploading to Database: {e}")

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
                # 🛠️ Convert Stock Column to Numeric safely
                if 'Current_Stock_Units' in df.columns:
                    df['Current_Stock_Units'] = pd.to_numeric(df['Current_Stock_Units'], errors='coerce').fillna(0)
                
                store_desc_col = 'Store_Description' if 'Store_Description' in df.columns else 'Store'
                item_column = 'SKU_Description' if 'SKU_Description' in df.columns else 'SKU'

                # 🛠️ Flexible Warehouse Detection (Contains DCW1 or Kerawalapitiya)
                warehouse_mask = (
                    df[store_desc_col].astype(str).str.contains('DCW1|Kerawalapitiya', case=False, na=False) |
                    df.get('Plant', pd.Series()).astype(str).str.contains('DCW1', case=False, na=False)
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
                    else:
                        st.info("Outlets කිසිවක් හමු වූයේ නැත.")

                # ================= 3. ZERO STOCK REPORT =================
                elif main_menu == "⚠️ Zero Stock Report":
                    st.subheader(f"📋 Zero Stock Outlets ({selected_batch})")
                    sub_tab1, sub_tab2 = st.tabs(["🥛 Dairies", "🍚 Rice"])

                    def render_zero_stock_section(category_name):
                        cat_df = outlets_df[outlets_df['Category'] == category_name]
                        cat_items = sorted(cat_df[item_column].dropna().unique())
                        
                        if not cat_items:
                            st.info(f"No items found in {category_name} category.")
                            return

                        selected_zero_item = st.selectbox(f"📦 Select {category_name} Item", cat_items, key=f"zero_{category_name}")
                        zero_df = cat_df[(cat_df[item_column] == selected_zero_item) & (cat_df['Current_Stock_Units'] <= 0)]

                        if not zero_df.empty:
                            st.error(f"🚨 Outlets {len(zero_df)} ක මේ Item එක Zero Stock වී ඇත!")
                            display_cols = [store_desc_col, 'SKU', 'Current_Stock_Units', 'Material_Status_Desc']
                            available_disp = [c for c in display_cols if c in zero_df.columns]
                            report_df = zero_df[available_disp].reset_index(drop=True)

                            st.dataframe(report_df, use_container_width=True)
                        else:
                            st.success(f"✅ නියමයි! මේ {category_name} Item එක හැම Outlet එකකම Stock තියෙනවා.")

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
                        st.warning("⚠️ Warehouse (DCW1) එකට අදාළ Records හමු වූයේ නැත. (Store Description එකේ DCW1 හෝ Kerawalapitiya තිබේදැයි බලන්න)")

    except Exception as e:
        st.error(f"Error fetching data from database: {e}")
