import streamlit as st
import pandas as pd
import os

# Page setup (Mobile friendly)
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🛒 Keells Stock Tracker")

# Excel file එකේ නම
FILE_NAME = "App.xlsx"

# File එක අන්තිමට වෙනස් වුණු වෙලාව අනුව cache එක auto invalidate කරනවා
def get_file_mtime(filepath):
    try:
        return os.path.getmtime(filepath)
    except:
        return 0

# Load Data (TTL සහ modification time එක පාවිච්චි කරලා ලෙඩේ නැති කරනවා)
@st.cache_data(ttl=600, hash_funcs={float: lambda x: int(x)})
def load_data(mtime):
    df = pd.read_excel(FILE_NAME)
    if 'SKU' in df.columns:
        df['SKU'] = df['SKU'].astype(str).str.replace(r'\.0$', '', regex=True)
    return df

try:
    # File modification time එක ගන්නවා
    file_mtime = get_file_mtime(FILE_NAME)
    df = load_data(file_mtime)

    # 1. Select Outlet
    outlets = sorted(df['Store Description'].unique())
    selected_outlet = st.selectbox("📍 Select Outlet / Store", outlets)

    # Filter data for selected outlet
    outlet_df = df[df['Store Description'] == selected_outlet]

    # --- ITEM SELECTION BY SKU DESCRIPTION ---
    item_column = 'SKU Description' if 'SKU Description' in df.columns else df.columns[0]
    
    items = sorted(outlet_df[item_column].dropna().unique())
    selected_item = st.selectbox("📦 Select Item", items)

    # Filter data for selected item in that outlet
    item_details = outlet_df[outlet_df[item_column] == selected_item].iloc[0]

    # --- DETAIL VIEW ---
    st.markdown("---")
    st.subheader(f"🔹 {selected_item}")
    
    # Styled Card View for SKU
    st.info(f"**SKU:** {item_details.get('SKU', 'N/A')}")
    
    # Grid Layout
    col1, col2 = st.columns(2)
    
    with col1:
        st.write(f"🏢 **Store Description:** {item_details.get('Store Description', 'N/A')}")
        st.write(f"📊 **Current Stock On Hand:** `{item_details.get('Current Stock On Hand Units', 0)}` Units")
        st.write(f"🔄 **Last Update Time:** {item_details.get('Last Update Date Time', 'N/A')}")

    with col2:
        st.write(f"⚙️ **Material Status:** {item_details.get('Material Status', 'N/A')}")
        st.write(f"📝 **Status Description:** {item_details.get('Material Status Description', 'N/A')}")
        st.write(f"🔑 **Dairy Key:** `{item_details.get('Dairy_Key', 'N/A')}`")

    # Sidebar එකේ පල්ලෙහායින් මැනුවල් Refresh බටන් එකක්
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Clear App Cache / Refresh"):
        st.cache_data.clear()
        st.rerun()

except Exception as e:
    st.error(f"Error loading data: {e}")
    st.info("කරුණාකර Excel file එකේ Column names සහ නම (App.xlsx) නිවැරදිදැයි පරීක්ෂා කරන්න.")
