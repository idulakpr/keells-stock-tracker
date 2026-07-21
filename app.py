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
        df['SKU'] = df['SKU'].astype(str).str.replace(r'\.0$', '', regex=True)
    return df

try:
    file_mtime = get_file_mtime(FILE_NAME)
    df = load_data(file_mtime)

    # --- TABS FOR NAVIGATION ---
    tab1, tab2 = st.tabs(["🔍 Outlet Stock Search", "⚠️ Zero Stock Report"])

    # ================= TAB 1: OUTLET SEARCH =================
    with tab1:
        outlets = sorted(df['Store Description'].dropna().unique())
        selected_outlet = st.selectbox("📍 Select Outlet / Store", outlets)

        outlet_df = df[df['Store Description'] == selected_outlet]

        item_column = 'SKU Description' if 'SKU Description' in df.columns else df.columns[0]
        items = sorted(outlet_df[item_column].dropna().unique())
        selected_item = st.selectbox("📦 Select Item", items)

        item_details = outlet_df[outlet_df[item_column] == selected_item].iloc[0]

        st.markdown("---")
        st.subheader(f"🔹 {selected_item}")
        
        st.info(f"**SKU:** {item_details.get('SKU', 'N/A')}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"🏢 **Store Description:** {item_details.get('Store Description', 'N/A')}")
            st.write(f"📊 **Current Stock On Hand:** `{item_details.get('Current Stock On Hand Units', 0)}` Units")
            st.write(f"🔄 **Last Update Time:** {item_details.get('Last Update Date Time', 'N/A')}")

        with col2:
            st.write(f"⚙️ **Material Status:** {item_details.get('Material Status', 'N/A')}")
            st.write(f"📝 **Status Description:** {item_details.get('Material Status Description', 'N/A')}")
            st.write(f"🔑 **Dairy Key:** `{item_details.get('Dairy_Key', 'N/A')}`")

    # ================= TAB 2: ZERO STOCK REPORT =================
    with tab2:
        st.subheader("📋 Item-wise Zero Stock Outlets")
        st.caption("තෝරන Item එක සඳහා Stock 0 තියෙන Outlets ලැයිස්තුව:")

        item_column = 'SKU Description' if 'SKU Description' in df.columns else df.columns[0]
        all_items = sorted(df[item_column].dropna().unique())
        
        selected_zero_item = st.selectbox("📦 Select Item for Zero Stock Check", all_items, key="zero_item")

        # Stock <= 0 තියෙන Outlets Filter කිරීම
        zero_df = df[(df[item_column] == selected_zero_item) & (df['Current Stock On Hand Units'] <= 0)]

        if not zero_df.empty:
            st.error(f"🚨 Outlets {len(zero_df)} ක මේ Item එක Zero Stock වී ඇත!")

            # පෙන්විය යුතු Columns ටික විතරක් තෝරා ගැනීම
            display_cols = ['Store Description', 'SKU', 'Current Stock On Hand Units', 'Material Status Description']
            valid_cols = [col for col in display_cols if col in zero_df.columns]
            
            report_df = zero_df[valid_cols].reset_index(drop=True)
            report_df.columns = [col.replace('Current Stock On Hand Units', 'Stock On Hand') for col in report_df.columns]

            # Table එක පෙන්වීම
            st.dataframe(report_df, use_container_width=True)

            # CSV Download Button එක
            csv = report_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Zero Stock Report (CSV)",
                data=csv,
                file_name=f"Zero_Stock_{selected_zero_item}.csv",
                mime="text/csv",
            )
        else:
            st.success("✅ නියමයි! මේ Item එක හැම Outlet එකකම Stock තියෙනවා (Zero Stock Outlets නැත).")

    # Sidebar Refresh
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Clear App Cache / Refresh"):
        st.cache_data.clear()
        st.rerun()

except Exception as e:
    st.error(f"Error loading data: {e}")
    st.info("කරුණාකර Excel file එකේ Column names නිවැරදිදැයි පරීක්ෂා කරන්න.")
