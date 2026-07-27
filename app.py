import streamlit as st
import pandas as pd
import os

# Page setup (Mobile friendly)
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🛒 Keells Stock Tracker")

FILE_NAME = "App.xlsx"

def get_file_mtime(filepath):
    try:
        return os.path.getmtime(filepath)
    except:
        return 0

@st.cache_data(ttl=600, hash_funcs={float: lambda x: int(x)})
def load_data(mtime):
    df = pd.read_excel(FILE_NAME)
    if 'SKU' in df.columns:
        df['SKU'] = df['SKU'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    return df

try:
    file_mtime = get_file_mtime(FILE_NAME)
    df = load_data(file_mtime)

    # Column identifications
    item_column = 'SKU Description' if 'SKU Description' in df.columns else df.columns[0]
    store_column = 'Store Description' if 'Store Description' in df.columns else 'Store'

    # --- CATEGORIZATION BY SKU CODE ---
    dairy_skus = ['115281', '115282', '115283', '5285', '44132', '126507', '128484', '120115']

    def categorize_by_sku(row):
        sku_val = str(row.get('SKU', '')).strip()
        if sku_val in dairy_skus:
            return 'Dairies'
        return 'Rice'

    df['Category'] = df.apply(categorize_by_sku, axis=1)

    # --- WAREHOUSE FILTER (DCW1) ---
    warehouse_mask = df[store_column].astype(str).str.strip().str.upper() == 'DCW1'
    
    warehouse_df = df[warehouse_mask]
    outlets_df = df[~warehouse_mask]

    # --- VERTICAL MENU SELECTION (පහළට එක යට එක) ---
    st.markdown("### 📌 Navigation Menu")
    main_menu = st.radio(
        "ඔයාට අවශ්‍ය Option එක තෝරන්න:",
        ["🔍 Outlet Stock Search", "⚠️ Zero Stock Report", "🏬 Warehouse Stock"],
        index=0
    )

    st.markdown("---")

    # ================= 1. OUTLET STOCK SEARCH =================
    if main_menu == "🔍 Outlet Stock Search":
        outlets = sorted(outlets_df[store_column].dropna().unique())
        selected_outlet = st.selectbox("📍 Select Outlet / Store", outlets)

        outlet_data = outlets_df[outlets_df[store_column] == selected_outlet]

        items = sorted(outlet_data[item_column].dropna().unique())
        selected_item = st.selectbox("📦 Select Item", items)

        item_details = outlet_data[outlet_data[item_column] == selected_item].iloc[0]

        st.markdown("---")
        st.subheader(f"🔹 {selected_item}")
        
        st.info(f"**SKU:** {item_details.get('SKU', 'N/A')}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"🏢 **Store Description:** {item_details.get(store_column, 'N/A')}")
            st.write(f"📊 **Current Stock On Hand:** `{item_details.get('Current Stock On Hand Units', 0)}` Units")
            st.write(f"🔄 **Last Update Time:** {item_details.get('Last Update Date Time', 'N/A')}")

        with col2:
            st.write(f"⚙️ **Material Status:** {item_details.get('Material Status', 'N/A')}")
            st.write(f"📝 **Status Description:** {item_details.get('Material Status Description', 'N/A')}")
            st.write(f"🔑 **Dairy Key:** `{item_details.get('Dairy_Key', 'N/A')}`")

    # ================= 2. ZERO STOCK REPORT =================
    elif main_menu == "⚠️ Zero Stock Report":
        st.subheader("📋 Item-wise Zero Stock Outlets")
        
        sub_tab1, sub_tab2 = st.tabs(["🥛 Dairies", "🍚 Rice"])

        def render_zero_stock_section(category_name):
            cat_df = outlets_df[outlets_df['Category'] == category_name]
            cat_items = sorted(cat_df[item_column].dropna().unique())
            
            if not cat_items:
                st.info(f"No items found in {category_name} category.")
                return

            selected_zero_item = st.selectbox(f"📦 Select {category_name} Item", cat_items, key=f"zero_{category_name}")

            zero_df = cat_df[(cat_df[item_column] == selected_zero_item) & (cat_df['Current Stock On Hand Units'] <= 0)]

            if not zero_df.empty:
                st.error(f"🚨 Outlets {len(zero_df)} ක මේ Item එක Zero Stock වී ඇත!")

                display_cols = [store_column, 'SKU', 'Current Stock On Hand Units', 'Material Status Description']
                valid_cols = [col for col in display_cols if col in zero_df.columns]
                
                report_df = zero_df[valid_cols].reset_index(drop=True)
                report_df.columns = [col.replace('Current Stock On Hand Units', 'Stock On Hand') for col in report_df.columns]

                st.dataframe(report_df, use_container_width=True)

                csv = report_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Download {category_name} Zero Stock Report (CSV)",
                    data=csv,
                    file_name=f"Zero_Stock_{category_name}_{selected_zero_item}.csv",
                    mime="text/csv",
                    key=f"dl_{category_name}"
                )
            else:
                st.success(f"✅ නියමයි! මේ {category_name} Item එක හැම Outlet එකකම Stock තියෙනවා.")

        with sub_tab1:
            render_zero_stock_section("Dairies")

        with sub_tab2:
            render_zero_stock_section("Rice")

    # ================= 3. WAREHOUSE STOCK =================
    elif main_menu == "🏬 Warehouse Stock":
        st.subheader("🏬 Warehouse Stock (DCW1)")
        st.caption("Warehouse (DCW1) එකේ දැනට තියෙන සම්පූර්ණ Stock මට්ටම්:")

        if not warehouse_df.empty:
            wh_display_cols = ['SKU', item_column, 'Category', 'Current Stock On Hand Units', 'Material Status Description']
            valid_wh_cols = [col for col in wh_display_cols if col in warehouse_df.columns]

            clean_wh_df = warehouse_df[valid_wh_cols].reset_index(drop=True)
            clean_wh_df.columns = [col.replace('Current Stock On Hand Units', 'Stock On Hand') for col in clean_wh_df.columns]

            st.dataframe(clean_wh_df, use_container_width=True)

            csv_wh = clean_wh_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Warehouse Stock Report (CSV)",
                data=csv_wh,
                file_name="Warehouse_Stock_DCW1.csv",
                mime="text/csv"
            )
        else:
            st.warning("⚠️ Warehouse (DCW1) එකට අදාළ Data හමු වූයේ නැත.")

    # Sidebar Refresh
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Clear App Cache / Refresh"):
        st.cache_data.clear()
        st.rerun()

except Exception as e:
    st.error(f"Error loading data: {e}")
    st.info("කරුණාකර Excel file එකේ Column names නිවැරදිදැයි පරීක්ෂා කරන්න.")
