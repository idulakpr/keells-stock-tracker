import streamlit as st
import pandas as pd

# Mobile view එකට සෙට් වෙන්න layout එක හදාගැනීම
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🌾 Keells Live Stock Tracker 🥛")
st.write("Real-time Rice & Dairy Stock Levels")

# 1. Excel File එක කියවීම (App.xlsx ඔයා upload කරපු file එකමයි)
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("App.xlsx")
        # Columns වල තියෙන හිස්තැන් අයින් කිරීම
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error("Excel file එක කියවන්න බැහැ. 'App.xlsx' file එක නිවැරදිව තියෙනවද බලන්න.")
        return None

df = load_data()

if df is not None:
    # 2. Dropdown එකක් මඟින් Outlet එක තෝරාගැනීම
    outlets = sorted(df['Store Description'].unique())
    selected_outlet = st.selectbox("📍 Select Keells Outlet:", outlets)
    
    if selected_outlet:
        # තෝරාගත් Outlet එකට අදාළ data විතරක් filter කිරීම
        outlet_df = df[df['Store Description'] == selected_outlet]
        
        # Last update වෙලාව පෙන්වීම
        if not outlet_df.empty:
            last_update = outlet_df['Last Update Date Time'].iloc[0]
            st.info(f"🕒 Last Updated: {last_update}")
        
        # 3. 🥛 DAIRY ITEMS සෙක්ෂන් එක
        st.subheader("🥛 Dairy Items")
        # Dairy_Key එක හිස් නොවන (NaN නොවන) සහ Stock එකක් තියෙන ඒවා filter කිරීම
        dairy_df = outlet_df[outlet_df['Dairy_Key'].notna()]
        
        if not dairy_df.empty:
            for idx, row in dairy_df.iterrows():
                stock = int(row['Current Stock On Hand Units'])
                # Stock එක 0 නම් රතු පාටින්, නැත්නම් කොළ පාටින් පෙන්වන්න
                color = "green" if stock > 0 else "red"
                st.markdown(f"**{row['SKU Description']}**  \nStock: :{color}[{stock} Units]")
                st.divider()
        else:
            st.write("No Dairy items found for this outlet.")
            
        # 4. 🌾 RICE ITEMS සෙක්ෂන් එක
        st.subheader("🌾 Rice Items")
        # Rice_Key එක හිස් නොවන (NaN නොවන) ඒවා filter කිරීම
        rice_df = outlet_df[outlet_df['Rice_Key'].notna()]
        
        if not rice_df.empty:
            for idx, row in rice_df.iterrows():
                stock = int(row['Current Stock On Hand Units'])
                color = "green" if stock > 0 else "red"
                st.markdown(f"**{row['SKU Description']}**  \nStock: :{color}[{stock} Units]")
                st.divider()
        else:
            st.write("No Rice items found for this outlet.")
