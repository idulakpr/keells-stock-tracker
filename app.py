import streamlit as st
import pandas as pd
import plotly.express as px
from supabase import create_client, Client
import datetime

# Page setup (Mobile friendly)
st.set_page_config(page_title="Keells Stock Tracker", layout="centered")

st.title("🛒 Keells Stock Tracker")

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

# --- HELPER FUNCTION: SUPABASE PAGINATION FOR SINGLE BATCH FETCH ---
@st.cache_data(ttl=300, show_spinner=False)
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
        
    df = pd.DataFrame(all_rows)
    
    if not df.empty and 'Store' in df.columns and 'SKU' in df.columns:
        df = df.drop_duplicates(subset=['Store', 'SKU'], keep='last')
        
    return df

# --- HELPER FUNCTION: FAST VIEW-BASED UNIQUE BADGES FETCH ---
def get_unique_badges():
    try:
        response = supabase.table('unique_badges_view').select('Uploaded_At').execute()
        if response.data:
            badges = [row['Uploaded_At'] for row in response.data if row.get('Uploaded_At')]
            return sorted(badges, reverse=True)
        return []
    except Exception as e:
        try:
            res = supabase.table('stock_history').select('Uploaded_At').execute()
            if res.data:
                badges = list(set([r['Uploaded_At'] for r in res.data if r.get('Uploaded_At')]))
                return sorted(badges, reverse=True)
        except:
            pass
        return []

# --- NAVIGATION MENU ---
st.markdown("### 📌 Navigation Menu")

if st.session_state.is_admin:
    st.sidebar.success("🔓 Logged in as Admin")
    if st.sidebar.button("🔒 Logout Admin"):
        st.session_state.is_admin = False
        st.rerun()

main_menu = st.radio(
    "ඔයාට අවශ්‍ය Option එක තෝරන්න:",
    [
        "🔍 Outlet Stock Search (Latest)", 
        "⚠️ Zero Stock Report (Latest)", 
        "🏬 Warehouse Stock (Latest)",
        "📈 Historical OOS Trend Analysis",
        "📤 Upload New Stock (Admin)", 
        "🗑️ Manage / Delete Uploads (Admin)"
    ],
    index=0
)

st.markdown("---")

timestamps = get_unique_badges()
latest_badge = timestamps[0] if timestamps else None

# ================= 1. UPLOAD NEW STOCK (PROTECTED) =================
if main_menu == "📤 Upload New Stock (Admin)":
    st.subheader("📤 Upload Daily Excel File to Database")
    
    if check_admin_password():
        st.caption("මෙතැනින් ඔයාගේ Stock Date එක සහ Batch එක තෝරලා Excel File එක Upload කරන්න.")

        col_date, col_slot = st.columns(2)
        with col_date:
            upload_date = st.date_input("📅 Stock Data අදාළ දිනය (Date):", datetime.date.today())
        
        with col_slot:
            batch_num = st.selectbox(
                "🔢 අද දවසේ කීවෙනි Upload එකද? (Batch Number):", 
                ["Batch 1", "Batch 2", "Batch 3", "Batch 4", "Batch 5"]
            )
        
        custom_note = st.text_input("📝 වෙනත් සටහනක් (Optional Note - e.g. Evening Update):", "")

        date_str = upload_date.strftime("%Y-%m-%d")
        if custom_note.strip():
            badge_name = f"{date_str} - {batch_num} ({custom_note.strip()})"
        else:
            badge_name = f"{date_str} - {batch_num}"

        st.info(f"🏷️ **මෙම File එක Save වන Badge එක:** `{badge_name}`")

        overwrite_flag = badge_name in timestamps
        if overwrite_flag:
            st.warning(f"⚠️ **`{badge_name}`** කියන Badge එක දැනටමත් Database එකේ තියෙනවා! Upload කළහොත් පැරණි Data Replace වෙනවා.")

        uploaded_file = st.file_uploader("Choose App.xlsx file", type=["xlsx", "xls"])

        if uploaded_file is not None:
            if st.button("🚀 Save & Memorize to Database"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    if overwrite_flag:
                        status_text.text("🔄 පැරණි Duplicate Records අයින් කරමින් පවතී...")
                        supabase.table('stock_history').delete().eq('Uploaded_At', badge_name).execute()

                    status_text.text("📖 Excel File එක කියවමින් පවතී...")
                    df = pd.read_excel(uploaded_file)

                    rename_dict = {
                        'SKU Description': 'SKU_Description',
                        'Store Description': 'Store_Description',
                        'Current Stock On Hand Units': 'Current_Stock_Units',
                        'Material Status Description': 'Material_Status_Desc',
                        'Last Update Date Time': 'Last_Update_Time'
                    }
                    df = df.rename(columns=rename_dict)

                    if 'SKU' in df.columns:
                        df['SKU'] = df['SKU'].astype(str).apply(lambda x: str(x).split('.')[0].strip())

                    df['Category'] = df['SKU'].apply(categorize_by_sku)

                    if 'Current_Stock_Units' in df.columns:
                        df['Current_Stock_Units'] = pd.to_numeric(df['Current_Stock_Units'], errors='coerce').fillna(0)

                    df['Uploaded_At'] = badge_name

                    valid_db_columns = [
                        'Uploaded_At', 'Store', 'Store_Description', 'SKU', 
                        'SKU_Description', 'Category', 'Current_Stock_Units', 
                        'Material_Status_Desc', 'Last_Update_Time'
                    ]
                    
                    cols_to_keep = [c for c in valid_db_columns if c in df.columns]
                    df_upload = df[cols_to_keep].copy()
                    df_upload = df_upload.where(pd.notnull(df_upload), None)

                    records = df_upload.to_dict(orient='records')
                    total_records = len(records)
                    
                    status_text.text(f"⬆️ Database එකට Data Upload වෙමින් පවතී... (Total Rows: {total_records})")

                    chunk_size = 200
                    for i in range(0, total_records, chunk_size):
                        chunk = records[i:i + chunk_size]
                        supabase.table('stock_history').insert(chunk).execute()
                        
                        progress = min((i + chunk_size) / total_records, 1.0)
                        progress_bar.progress(progress)

                    status_text.empty()
                    progress_bar.empty()
                    
                    st.cache_data.clear()
                    st.success(f"✅ Data successfully Memorized under Badge: '{badge_name}'!")
                    st.balloons()
                    st.rerun()
                    
                except Exception as e:
                    status_text.empty()
                    progress_bar.empty()
                    st.error(f"❌ Upload එක අසාර්ථක විය! Error Message: {e}")

# ================= 2. MANAGE / DELETE UPLOADED FILES (PROTECTED) =================
elif main_menu == "🗑️ Manage / Delete Uploads (Admin)":
    st.subheader("🗑️ Upload කරපු Excel Batches අයින් කිරීම")
    
    if check_admin_password():
        if not timestamps:
            st.info("ℹ️ Database එකේ කිසිම Data එකක් නැත.")
        else:
            selected_delete_batch = st.selectbox("❌ Delete කරන්න අවශ්‍ය Upload Batch / Badge එක තෝරන්න:", timestamps)

            st.warning(f"⚠️ ඔබ තෝරාගත් Badge එක (`{selected_delete_batch}`) ස්ථිරවම Database එකෙන් Delete වනු ඇත.")
            
            if st.button("🔴 Delete Selected Batch"):
                with st.spinner("Deleting Batch from Database..."):
                    try:
                        supabase.table('stock_history').delete().eq('Uploaded_At', selected_delete_batch).execute()
                        st.cache_data.clear()
                        st.success(f"✅ Badge '{selected_delete_batch}' සාර්ථකව Delete කරන ලදී!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error deleting data: {e}")

# ================= 3. HISTORICAL OOS TREND ANALYSIS (UPDATED WITH CATEGORIES) =================
elif main_menu == "📈 Historical OOS Trend Analysis":
    st.subheader("📈 Historical OOS Trend Analysis")
    
    if not timestamps:
        st.warning("⚠️ Database එකේ කිසිම Data එකක් නැත.")
    else:
        st.caption("කාල පරාසයක් සහ Item හෝ Outlet එකක් තෝරා OOS Trend එක බලන්න:")

        # Date Range Selection
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("📅 Start Date", datetime.date.today() - datetime.timedelta(days=7))
        with col2:
            end_date = st.date_input("📅 End Date", datetime.date.today())

        if start_date > end_date:
            st.error("❌ Start Date එක End Date එකට වඩා වැඩි විය නොහැක!")
        else:
            # Filter Badges by Date Range
            filtered_badges = []
            for b in timestamps:
                try:
                    b_date_str = b.split(' - ')[0].strip()
                    b_date = datetime.datetime.strptime(b_date_str, "%Y-%m-%d").date()
                    if start_date <= b_date <= end_date:
                        filtered_badges.append(b)
                except:
                    pass

            filtered_badges = sorted(filtered_badges)

            if not filtered_badges:
                st.warning("⚠️ ඔබ තෝරාගත් Date Range එක ඇතුළත කිසිදු Stock Batch එකක් හමු නොවීය.")
            else:
                st.success(f"🔍 Batches **{len(filtered_badges)}** ක් හමු විය.")

                # Analysis Mode Selection
                analysis_mode = st.radio(
                    "🔍 Analysis Type එක තෝරන්න:",
                    ["📦 Item-wise Analysis (Outlets count over time)", "🏬 Outlet-wise Analysis (Items count over time)"],
                    horizontal=True
                )

                sample_df = fetch_all_batch_data(filtered_badges[-1])
                item_col = 'SKU_Description' if 'SKU_Description' in sample_df.columns else 'SKU'
                store_c = 'Store_Description' if 'Store_Description' in sample_df.columns else 'Store'

                # --- MODE 1: ITEM WISE ---
                if "Item-wise" in analysis_mode:
                    # 📂 Category එක තෝරා ගැනීමට Filter එකක් එකතු කිරීම
                    selected_cat = st.radio("📂 Category එක තෝරන්න:", ["Dairies", "Rice"], horizontal=True, key="trend_cat_filter")
                    
                    # අදාළ Category එකට අදාළ Items පමණක් Filter කරගැනීම
                    cat_sample_df = sample_df[sample_df['Category'] == selected_cat]
                    all_items = sorted(cat_sample_df[item_col].dropna().unique().tolist())
                    
                    if not all_items:
                        st.warning(f"⚠️ {selected_cat} Category එකට අදාළ Items හමු වූයේ නැත.")
                    else:
                        selected_item = st.selectbox(f"📦 Filter by {selected_cat} Item:", all_items)

                        if st.button("🚀 Generate Item OOS Trend Graph"):
                            with st.spinner("Analyzing OOS history across selected dates..."):
                                trend_data = []

                                for badge in filtered_badges:
                                    batch_df = fetch_all_batch_data(badge)
                                    if not batch_df.empty:
                                        wh_mask = batch_df[store_c].astype(str).str.contains('DCW1|Kerawalapitiya', case=False, na=False)
                                        outlets_b_df = batch_df[~wh_mask]
                                        
                                        if 'Current_Stock_Units' in outlets_b_df.columns:
                                            outlets_b_df['Current_Stock_Units'] = pd.to_numeric(outlets_b_df['Current_Stock_Units'], errors='coerce').fillna(0)

                                        item_df = outlets_b_df[(outlets_b_df[item_col] == selected_item) & (outlets_b_df['Current_Stock_Units'] <= 0)]
                                        
                                        oos_outlets_list = sorted(item_df[store_c].dropna().unique().tolist())
                                        oos_count = len(oos_outlets_list)
                                        
                                        trend_data.append({
                                            "Batch / Date": badge,
                                            "OOS Outlets Count": oos_count,
                                            "OOS Outlets List": ", ".join(oos_outlets_list) if oos_outlets_list else "None"
                                        })

                                chart_df = pd.DataFrame(trend_data)

                                if not chart_df.empty:
                                    oos_days = len(chart_df[chart_df["OOS Outlets Count"] > 0])
                                    max_peak = chart_df["OOS Outlets Count"].max()

                                    st.markdown("---")
                                    m1, m2, m3 = st.columns(3)
                                    m1.metric("Checked Batches", f"{len(chart_df)}")
                                    m2.metric("OOS Occurred Batches", f"{oos_days} Batches", delta_color="inverse")
                                    m3.metric("Max OOS Outlets Peak", f"{max_peak} Outlets")

                                    # Line Graph with Markers
                                    st.subheader(f"📈 OOS Outlet Trend - {selected_item}")
                                    fig = px.line(
                                        chart_df, 
                                        x="Batch / Date", 
                                        y="OOS Outlets Count",
                                        text="OOS Outlets Count",
                                        markers=True,
                                        labels={"OOS Outlets Count": "Number of Outlets OOS", "Batch / Date": "Upload Batch / Date"}
                                    )
                                    fig.update_traces(
                                        line_color='#00a896', 
                                        line_width=4, 
                                        marker=dict(size=10, symbol='circle'),
                                        textposition='top center'
                                    )
                                    fig.update_layout(xaxis_tickangle=-45, yaxis=dict(zeroline=True))

                                    st.plotly_chart(fig, use_container_width=True)

                                    # Detailed Table with Outlet List
                                    st.subheader("📋 Batch-wise OOS Outlet Details")
                                    st.dataframe(chart_df, use_container_width=True)

                # --- MODE 2: OUTLET WISE ---
                else:
                    wh_mask_sample = sample_df[store_c].astype(str).str.contains('DCW1|Kerawalapitiya', case=False, na=False)
                    all_outlets = sorted(sample_df[~wh_mask_sample][store_c].dropna().unique().tolist())
                    selected_outlet = st.selectbox("🏬 Filter by Outlet / Store:", all_outlets)

                    if st.button("🚀 Generate Outlet OOS Trend Graph"):
                        with st.spinner("Analyzing OOS history across selected dates..."):
                            trend_data = []

                            for badge in filtered_badges:
                                batch_df = fetch_all_batch_data(badge)
                                if not batch_df.empty:
                                    if 'Current_Stock_Units' in batch_df.columns:
                                        batch_df['Current_Stock_Units'] = pd.to_numeric(batch_df['Current_Stock_Units'], errors='coerce').fillna(0)

                                    outlet_df = batch_df[(batch_df[store_c] == selected_outlet) & (batch_df['Current_Stock_Units'] <= 0)]
                                    
                                    oos_items_list = sorted(outlet_df[item_col].dropna().unique().tolist())
                                    oos_count = len(oos_items_list)
                                    
                                    trend_data.append({
                                        "Batch / Date": badge,
                                        "OOS Items Count": oos_count,
                                        "OOS Items List": ", ".join(oos_items_list) if oos_items_list else "None"
                                    })

                            chart_df = pd.DataFrame(trend_data)

                            if not chart_df.empty:
                                oos_days = len(chart_df[chart_df["OOS Items Count"] > 0])
                                max_peak = chart_df["OOS Items Count"].max()

                                st.markdown("---")
                                m1, m2, m3 = st.columns(3)
                                m1.metric("Checked Batches", f"{len(chart_df)}")
                                m2.metric("OOS Occurred Batches", f"{oos_days} Batches", delta_color="inverse")
                                m3.metric("Max OOS Items Peak", f"{max_peak} Items")

                                # Line Graph with Markers
                                st.subheader(f"📈 OOS Item Trend - {selected_outlet}")
                                fig = px.line(
                                    chart_df, 
                                    x="Batch / Date", 
                                    y="OOS Items Count",
                                    text="OOS Items Count",
                                    markers=True,
                                    labels={"OOS Items Count": "Number of Items OOS", "Batch / Date": "Upload Batch / Date"}
                                )
                                fig.update_traces(
                                    line_color='#e63946', 
                                    line_width=4, 
                                    marker=dict(size=10, symbol='circle'),
                                    textposition='top center'
                                )
                                fig.update_layout(xaxis_tickangle=-45, yaxis=dict(zeroline=True))

                                st.plotly_chart(fig, use_container_width=True)

                                # Detailed Table with Item List
                                st.subheader("📋 Batch-wise OOS Item Details")
                                st.dataframe(chart_df, use_container_width=True)

# ================= 4. FAST LATEST DATA RETRIEVAL FOR DAILY DASHBOARDS =================
else:
    if not latest_badge:
        st.warning("⚠️ Database එකේ කිසිම Data එකක් නෑ. කරුණාකර Admin හරහා Excel File එකක් Upload කරන්න.")
    else:
        st.caption(f"⚡ **Showing Latest Update:** `{latest_badge}`")

        with st.spinner("Fast loading latest stock data..."):
            df = fetch_all_batch_data(latest_badge)

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

            # --- OUTLET STOCK SEARCH (FAST) ---
            if main_menu == "🔍 Outlet Stock Search (Latest)":
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

            # --- ZERO STOCK REPORT (FAST) ---
            elif main_menu == "⚠️ Zero Stock Report (Latest)":
                st.subheader(f"📋 Zero Stock Outlets Report (Latest)")
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

            # --- WAREHOUSE STOCK (FAST) ---
            elif main_menu == "🏬 Warehouse Stock (Latest)":
                st.subheader("🏬 Warehouse Stock - DCW1 (Latest)")
                
                if not warehouse_df.empty:
                    wh_display_cols = ['SKU', item_column, 'Current_Stock_Units', 'Category']
                    available_wh = [c for c in warehouse_df.columns if c in wh_display_cols]
                    clean_wh_df = warehouse_df[available_wh].reset_index(drop=True)

                    st.dataframe(clean_wh_df, use_container_width=True)
                else:
                    st.warning("⚠️ Warehouse (DCW1) එකට අදාළ Records හමු වූයේ නැත.")
