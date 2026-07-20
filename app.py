import streamlit as st
import pandas as pd

# Page setup (Mobile friendly)
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🛒 Keells Stock Tracker")

# Load Data
@st.cache_data
def load_data():
    df = pd.read_excel("App.xlsx")
    # SKU එක හැමතිස්සෙම string එකක් විදිහට ගන්න (නැත්නම් .0 වැටෙන්න පුළුවන්)
    if 'SKU' in df.columns:
        df['SKU'] = df['SKU'].astype(str).str.replace(r'\.0$', '', regex=True)
    return df

try:
    df = load_data()

    # 1. Select Outlet
    outlets = sorted(df['Store Description'].unique())
    selected_outlet = st.selectbox("📍 Select Outlet / Store", outlets)

    # Filter data for selected outlet
    outlet_df = df[df['Store Description'] == selected_outlet]

    # --- ITEM SELECTION BY SKU DESCRIPTION ---
    # Excel එකේ column එක 'SKU Description' ද කියලා බලනවා, නැත්නම් තියෙන වෙන එකක් ගන්නවා
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

except Exception as e:
    st.error(f"Error loading data: {e}")
    st.info("කරුණාකර Excel file එකේ Column names නිවැරදිදැයි පරීක්ෂා කරන්න.")
