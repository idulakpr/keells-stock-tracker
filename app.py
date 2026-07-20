import streamlit as st
import pandas as pd

# Page setup (Mobile friendly)
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🛒 Keells Stock Tracker")

# Load Data
@st.cache_data
def load_data():
    # Excel file එක කියවීම
    df = pd.read_excel("App.xlsx")
    return df

try:
    df = load_data()

    # 1. Select Outlet
    outlets = sorted(df['Store Description'].unique())
    selected_outlet = st.selectbox("📍 Select Outlet / Store", outlets)

    # Filter data for selected outlet
    outlet_df = df[df['Store Description'] == selected_outlet]

    # 2. Select Item (Assuming column name is 'Item Description')
    # *Note: standard column name එක වෙනස් නම් පල්ලෙහා 'Item Description' කෑල්ල මාරු කරන්න.
    item_column = 'Item Description' if 'Item Description' in df.columns else df.columns[0] 
    
    items = sorted(outlet_df[item_column].unique())
    selected_item = st.selectbox("📦 Select Item", items)

    # Filter data for selected item in that outlet
    item_details = outlet_df[outlet_df[item_column] == selected_item].iloc[0]

    # --- DETAIL VIEW (ඔයා ඉල්ලපු විදිහට ලස්සනට පෙන්වීම) ---
    st.markdown("---")
    st.subheader(f"🔹 {selected_item}")
    
    # Styled Card View
    st.info(f"**SKU:** {item_details.get('SKU', 'N/A')}")
    
    # Grid එකක් වගේ පෙන්වන්න Columns 2කට බෙදමු
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
