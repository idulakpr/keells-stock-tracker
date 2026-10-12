import streamlit as st
import pandas as pd
import plotly.express as px
from supabase import create_client, Client
import datetime
import io
import hashlib
import json
import math
import re
import io


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Modern Trade Stock & Sales Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Modern Trade Stock & Sales Dashboard")
st.caption("Stock Monitoring | OOS Analysis | All-Channel Sales Performance")

# =========================================================
# SUPABASE CONNECTION
# =========================================================

try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")
except Exception as e:
    st.error(f"Supabase connection settings check කරන්න: {e}")
    st.stop()

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False


def check_admin_password():
    if st.session_state.is_admin:
        return True

    st.info("🔒 Admin access අවශ්‍යයි.")
    password = st.text_input("Admin Password", type="password")

    if st.button("Unlock Admin Features"):
        if password == ADMIN_PASSWORD:
            st.session_state.is_admin = True
            st.rerun()
        else:
            st.error("Incorrect password.")

    return False


# =========================================================
# COMMON HELPERS
# =========================================================

DAIRY_SKUS = [
    "115281", "115282", "115283", "5285",
    "44132", "126507", "128484", "120115"
]


def categorize_by_sku(sku):
    if pd.isna(sku):
        return "Rice"

    sku = str(sku).split(".")[0].strip()
    return "Dairies" if sku in DAIRY_SKUS else "Rice"


def dataframe_to_excel(df, sheet_name="Report"):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])

    return output.getvalue()


def multi_sheet_excel(sheets):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])

    return output.getvalue()


def fetch_all_rows(table_name, page_size=1000):
    rows = []
    start = 0

    while True:
        response = (
            supabase.table(table_name)
            .select("*")
            .range(start, start + page_size - 1)
            .execute()
        )

        batch = response.data or []
        rows.extend(batch)

        if len(batch) < page_size:
            break

        start += page_size

    return pd.DataFrame(rows)


def make_json_safe(value):
    """
    Convert pandas/Python values into JSON-compatible values.
    Handles date, datetime, NaN, infinity and NumPy scalar values.
    """
    if value is None:
        return None

    if not pd.api.types.is_scalar(value):
        return str(value)

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()

    if hasattr(value, "item"):
        value = value.item()

    if isinstance(value, float) and not math.isfinite(value):
        return None

    return value


def make_row_hash(row):
    """
    Create a stable hash from the business fields of a sales line.
    Identical lines are ignored on subsequent uploads.
    """
    fields = [
        "billing_date",
        "sap_invoice_number",
        "sales_group",
        "customer_code",
        "sold_to_party",
        "material_group",
        "material_code",
        "material_description",
        "invoiced_quantity",
        "unit",
        "net_weight",
        "unit_of_weight",
        "net_value_usd",
        "po_number",
        "special_instructions",
    ]

    values = {}

    for field in fields:
        values[field] = make_json_safe(row.get(field))

    payload = json.dumps(
        values,
        sort_keys=True,
        ensure_ascii=False
    )

    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def get_unique_badges():
    try:
        response = (
            supabase.table("unique_badges_view")
            .select("Uploaded_At")
            .execute()
        )

        if response.data:
            return sorted(
                {
                    row["Uploaded_At"]
                    for row in response.data
                    if row.get("Uploaded_At")
                },
                reverse=True
            )
    except Exception:
        pass

    try:
        df = fetch_all_rows("stock_history")

        if not df.empty and "Uploaded_At" in df.columns:
            return sorted(
                df["Uploaded_At"].dropna().unique().tolist(),
                reverse=True
            )
    except Exception:
        pass

    return []


@st.cache_data(ttl=300, show_spinner=False)
def fetch_all_batch_data(selected_batch):
    rows = []
    page_size = 1000
    start = 0

    while True:
        response = (
            supabase.table("stock_history")
            .select("*")
            .eq("Uploaded_At", selected_batch)
            .range(start, start + page_size - 1)
            .execute()
        )

        batch = response.data or []
        rows.extend(batch)

        if len(batch) < page_size:
            break

        start += page_size

    df = pd.DataFrame(rows)

    if not df.empty and {"Store", "SKU"}.issubset(df.columns):
        df = df.drop_duplicates(
            subset=["Store", "SKU"],
            keep="last"
        )

    return df


# =========================================================
# NAVIGATION
# =========================================================

st.sidebar.title("📌 Navigation")

if st.session_state.is_admin:
    st.sidebar.success("Admin logged in")

    if st.sidebar.button("🔒 Logout Admin"):
        st.session_state.is_admin = False
        st.rerun()

main_menu = st.sidebar.radio(
    "Select a section",
    [
        "📊 All-Channel Sales Analysis",
        "📤 Upload Sales Report (Admin)",
        "📄 Arpico PO PDF Upload (Admin)",
        "📊 Arpico PO Analysis",
        "🔍 Outlet Stock Search (Latest)",
        "⚠️ Zero Stock Report (Latest)",
        "🏬 Warehouse Stock (Latest)",
        "📈 Historical OOS Trend Analysis",
        "📤 Upload New Stock (Admin)",
        "🗑️ Manage / Delete Uploads (Admin)",
        "🗑️ Manage Sales Uploads (Admin)",
    ]
)

st.markdown("---")

stock_badges = get_unique_badges()
latest_badge = stock_badges[0] if stock_badges else None


# =========================================================
# 1. SALES REPORT UPLOAD
# =========================================================

if main_menu == "📤 Upload Sales Report (Admin)":

    st.header("📤 Upload Modern Trade Sales Report")

    if check_admin_password():

        batch_note = st.text_input(
            "Upload description",
            value=datetime.date.today().isoformat()
        )

        uploaded_sales_file = st.file_uploader(
            "Choose Sales Excel File",
            type=["xlsx", "xls"],
            key="sales_upload"
        )

        if uploaded_sales_file is not None:

            try:
                sales_df = pd.read_excel(uploaded_sales_file)
                sales_df.columns = [
                    str(col).strip() for col in sales_df.columns
                ]

                st.subheader("Sales Report Preview")
                st.dataframe(
                    sales_df.head(10),
                    use_container_width=True
                )

                if st.button(
                    "🚀 Upload Sales Data",
                    type="primary"
                ):

                    required_columns = [
                        "Billing Date",
                        "SAP Invoice Number",
                        "Sales Group",
                        "Customer",
                        "Sold-to-Party",
                        "Material Group",
                        "Material Code",
                        "Material Description",
                        "Invoiced Quantity",
                        "Unit",
                        "Net Weight",
                        "Unit of Weight",
                        "Net Value (USD)",
                        "PO number",
                        "Special Instructions",
                    ]

                    missing = [
                        col for col in required_columns
                        if col not in sales_df.columns
                    ]

                    if missing:
                        st.error(
                            "Missing Excel columns: "
                            + ", ".join(missing)
                        )

                    else:
                        rename_columns = {
                            "Billing Date": "billing_date",
                            "SAP Invoice Number": "sap_invoice_number",
                            "Sales Group": "sales_group",
                            "Customer": "customer_code",
                            "Sold-to-Party": "sold_to_party",
                            "Material Group": "material_group",
                            "Material Code": "material_code",
                            "Material Description": "material_description",
                            "Invoiced Quantity": "invoiced_quantity",
                            "Unit": "unit",
                            "Net Weight": "net_weight",
                            "Unit of Weight": "unit_of_weight",
                            "Net Value (USD)": "net_value_usd",
                            "PO number": "po_number",
                            "Special Instructions": "special_instructions",
                        }

                        df = sales_df.rename(
                            columns=rename_columns
                        ).copy()

                        # -------------------------------------------------
                        # DATE CONVERSION
                        # -------------------------------------------------
                        # Excel dates are parsed before hashing.
                        # Dates are converted to ISO strings before upload.

                        df["billing_date"] = pd.to_datetime(
                            df["billing_date"],
                            errors="coerce"
                        ).dt.date

                        invalid_dates = int(
                            df["billing_date"].isna().sum()
                        )

                        df = df.dropna(
                            subset=["billing_date"]
                        ).copy()

                        # -------------------------------------------------
                        # NUMERIC CONVERSION
                        # -------------------------------------------------

                        for col in [
                            "invoiced_quantity",
                            "net_weight",
                            "net_value_usd",
                        ]:
                            df[col] = pd.to_numeric(
                                df[col],
                                errors="coerce"
                            )

                        # -------------------------------------------------
                        # TEXT CONVERSION
                        # -------------------------------------------------

                        text_columns = [
                            "sap_invoice_number",
                            "sales_group",
                            "customer_code",
                            "sold_to_party",
                            "material_group",
                            "material_code",
                            "material_description",
                            "unit",
                            "unit_of_weight",
                            "po_number",
                            "special_instructions",
                        ]

                        for col in text_columns:
                            df[col] = df[col].apply(
                                lambda value: (
                                    None
                                    if pd.isna(value)
                                    else str(value).strip()
                                )
                            )

                        # -------------------------------------------------
                        # BATCH ID
                        # -------------------------------------------------

                        batch_timestamp = datetime.datetime.now().strftime(
                            "%Y%m%d_%H%M%S"
                        )

                        safe_note = (
                            batch_note.strip()
                            or "SalesUpload"
                        )

                        upload_batch = (
                            f"{batch_timestamp}_{safe_note}"
                        )

                        df["upload_batch"] = upload_batch

                        # Hash is created while billing_date is still
                        # a Python date. make_row_hash safely converts it.
                        df["row_hash"] = df.apply(
                            make_row_hash,
                            axis=1
                        )

                        # Remove duplicates within this Excel file.
                        df = df.drop_duplicates(
                            subset=["row_hash"],
                            keep="first"
                        )

                        db_columns = [
                            "upload_batch",
                            "row_hash",
                            "billing_date",
                            "sap_invoice_number",
                            "sales_group",
                            "customer_code",
                            "sold_to_party",
                            "material_group",
                            "material_code",
                            "material_description",
                            "invoiced_quantity",
                            "unit",
                            "net_weight",
                            "unit_of_weight",
                            "net_value_usd",
                            "po_number",
                            "special_instructions",
                        ]

                        df = df[db_columns]

                        records = df.to_dict(
                            orient="records"
                        )

                        # -------------------------------------------------
                        # CRITICAL FIX:
                        # Convert EVERY field to a JSON-safe value.
                        # This prevents "date is not JSON serializable".
                        # -------------------------------------------------

                        safe_records = []

                        for record in records:
                            safe_record = {
                                key: make_json_safe(value)
                                for key, value in record.items()
                            }

                            safe_records.append(safe_record)

                        if not safe_records:
                            st.warning(
                                "No valid sales rows found."
                            )

                        else:
                            progress = st.progress(0)
                            status = st.empty()

                            try:
                                chunk_size = 200
                                total_rows = len(safe_records)

                                for i in range(
                                    0,
                                    total_rows,
                                    chunk_size
                                ):
                                    chunk = safe_records[
                                        i:i + chunk_size
                                    ]

                                    supabase.table(
                                        "sales_history"
                                    ).upsert(
                                        chunk,
                                        on_conflict="row_hash",
                                        ignore_duplicates=True
                                    ).execute()

                                    completed = min(
                                        i + len(chunk),
                                        total_rows
                                    )

                                    progress.progress(
                                        completed / total_rows
                                    )

                                    status.write(
                                        f"Processed {completed:,} "
                                        f"of {total_rows:,} rows"
                                    )

                                st.cache_data.clear()

                                st.success(
                                    "Sales upload completed successfully!"
                                )

                                st.write(
                                    f"**Upload batch:** {upload_batch}"
                                )

                                st.write(
                                    f"**Unique rows processed:** "
                                    f"{total_rows:,}"
                                )

                                if invalid_dates:
                                    st.warning(
                                        f"{invalid_dates} rows with "
                                        "invalid billing dates were skipped."
                                    )

                                st.info(
                                    "Sales values are displayed as LKR "
                                    "without exchange-rate conversion. "
                                    "Existing identical rows are ignored."
                                )

                   
                            
                            except Exception as e:
                                st.error(
                                    f"Sales upload failed: {e}"
                                )

                                st.caption(
                                    "Check the sales_history table schema, "
                                    "database permissions and column types."
                                )

            except Exception as e:
                st.error(
                    f"Could not read the Excel file: {e}"
                )


# =========================================================
# 2. ALL-CHANNEL SALES ANALYSIS
# =========================================================
elif main_menu == "📊 All-Channel Sales Analysis":


    st.header("📊 All-Channel Modern Trade Sales Analysis")

    try:
        sales_df = fetch_all_rows("sales_history")

    except Exception as e:
        st.error(f"Could not retrieve sales data: {e}")
        sales_df = pd.DataFrame()


    if sales_df.empty:

        st.warning(
            "No sales data found. Upload a sales report first."
        )

    else:

        sales_df["billing_date"] = pd.to_datetime(
            sales_df["billing_date"],
            errors="coerce"
        )

        for col in [
            "net_value_usd",
            "invoiced_quantity",
            "net_weight",
        ]:
            sales_df[col] = pd.to_numeric(
                sales_df[col],
                errors="coerce"
            ).fillna(0)

        sales_df = sales_df.dropna(
            subset=["billing_date"]
        )

        min_date = sales_df["billing_date"].min().date()
        max_date = sales_df["billing_date"].max().date()

        st.subheader("Filters")

        f1, f2 = st.columns(2)

        with f1:
            start_date = st.date_input(
                "Start Date",
                value=min_date,
                min_value=min_date,
                max_value=max_date,
                key="sales_start"
            )

        with f2:
            end_date = st.date_input(
                "End Date",
                value=max_date,
                min_value=min_date,
                max_value=max_date,
                key="sales_end"
            )

        if start_date > end_date:
            st.error("Start Date cannot be after End Date.")
            st.stop()

        filtered = sales_df[
            (sales_df["billing_date"].dt.date >= start_date)
            & (sales_df["billing_date"].dt.date <= end_date)
        ].copy()

        channels = sorted(
            filtered["sold_to_party"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        items = sorted(
            filtered["material_description"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        c1, c2 = st.columns(2)

        with c1:
            selected_channels = st.multiselect(
                "Select Channel(s)",
                channels,
                default=channels
            )

        with c2:
            selected_items = st.multiselect(
                "Select Item(s)",
                items,
                default=items
            )

        filtered = filtered[
            filtered["sold_to_party"].astype(str).isin(
                selected_channels
            )
            & filtered["material_description"].astype(str).isin(
                selected_items
            )
        ].copy()

        total_sales = filtered["net_value_usd"].sum()
        total_weight = filtered["net_weight"].sum()
        invoice_count = filtered["sap_invoice_number"].nunique()
        channel_count = filtered["sold_to_party"].nunique()

        st.markdown("---")
        st.subheader("Performance Overview")

        m1, m2, m3, m4 = st.columns(4)

        m1.metric(
            "Total Sales Value",
            f"LKR {total_sales:,.2f}"
        )

        m2.metric(
            "Total Net Weight",
            f"{total_weight:,.2f} kg"
        )

        m3.metric(
            "Unique Invoices",
            f"{invoice_count:,}"
        )

        m4.metric(
            "Active Channels",
            f"{channel_count:,}"
        )

        st.caption(
            "Currency note: the source column is named 'Net Value (USD)', "
            "but values are displayed as LKR as requested. "
            "No currency conversion is applied."
        )

        if filtered.empty:

            st.warning("No sales data matches your filters.")

        else:

            # DAILY TREND
            st.markdown("---")
            st.subheader("📈 Daily Sales Trend")

            daily_source = filtered.copy()
            daily_source["Billing Date"] = (
                daily_source["billing_date"].dt.date
            )

            daily = (
                daily_source.groupby(
                    "Billing Date",
                    as_index=False
                )["net_value_usd"]
                .sum()
                .rename(
                    columns={
                        "net_value_usd": "Sales Value (LKR)"
                    }
                )
            )

            fig_daily = px.line(
                daily,
                x="Billing Date",
                y="Sales Value (LKR)",
                markers=True,
                title="Daily Sales Value (LKR)"
            )

            st.plotly_chart(
                fig_daily,
                use_container_width=True
            )

            # MONTHLY TREND
            st.subheader("📅 Monthly Sales Trend")

            monthly_source = filtered.copy()

            monthly_source["Month"] = (
                monthly_source["billing_date"]
                .dt.to_period("M")
                .astype(str)
            )

            monthly = (
                monthly_source.groupby(
                    "Month",
                    as_index=False
                )["net_value_usd"]
                .sum()
                .rename(
                    columns={
                        "net_value_usd": "Sales Value (LKR)"
                    }
                )
            )

            fig_monthly = px.bar(
                monthly,
                x="Month",
                y="Sales Value (LKR)",
                title="Monthly Sales Value (LKR)"
            )

            st.plotly_chart(
                fig_monthly,
                use_container_width=True
            )

            # CHANNEL SUMMARY
            st.markdown("---")
            st.subheader("🏬 Channel-wise Performance")

            channel_summary = (
                filtered.groupby(
                    "sold_to_party",
                    dropna=False
                )
                .agg(
                    Sales_Value_LKR=("net_value_usd", "sum"),
                    Net_Weight=("net_weight", "sum"),
                    Invoice_Count=("sap_invoice_number", "nunique"),
                    Item_Count=("material_code", "nunique"),
                )
                .reset_index()
                .rename(
                    columns={"sold_to_party": "Channel"}
                )
                .sort_values(
                    "Sales_Value_LKR",
                    ascending=False
                )
            )

            channel_summary["Sales_Value_LKR"] = (
                channel_summary["Sales_Value_LKR"].round(2)
            )

            channel_summary["Net_Weight"] = (
                channel_summary["Net_Weight"].round(2)
            )

            fig_channel = px.bar(
                channel_summary,
                x="Channel",
                y="Sales_Value_LKR",
                title="Channel-wise Sales Value (LKR)"
            )

            st.plotly_chart(
                fig_channel,
                use_container_width=True
            )

            st.dataframe(
                channel_summary,
                use_container_width=True,
                hide_index=True
            )

            # ITEM SUMMARY
            st.markdown("---")
            st.subheader("📦 Item-wise Performance")

            item_summary = (
                filtered.groupby(
                    [
                        "material_code",
                        "material_description",
                    ],
                    dropna=False
                )
                .agg(
                    Sales_Value_LKR=("net_value_usd", "sum"),
                    Invoiced_Quantity=("invoiced_quantity", "sum"),
                    Net_Weight=("net_weight", "sum"),
                    Invoice_Count=("sap_invoice_number", "nunique"),
                )
                .reset_index()
                .rename(
                    columns={
                        "material_code": "Material Code",
                        "material_description": "Material Description",
                    }
                )
                .sort_values(
                    "Sales_Value_LKR",
                    ascending=False
                )
            )

            item_summary["Sales_Value_LKR"] = (
                item_summary["Sales_Value_LKR"].round(2)
            )

            item_summary["Invoiced_Quantity"] = (
                item_summary["Invoiced_Quantity"].round(2)
            )

            item_summary["Net_Weight"] = (
                item_summary["Net_Weight"].round(2)
            )

            fig_items = px.bar(
                item_summary.head(20),
                x="Material Description",
                y="Sales_Value_LKR",
                title="Top 20 Items by Sales Value (LKR)"
            )

            fig_items.update_layout(
                xaxis_tickangle=-35
            )

            st.plotly_chart(
                fig_items,
                use_container_width=True
            )

            st.dataframe(
                item_summary,
                use_container_width=True,
                hide_index=True
            )

            # DETAILED RECORDS
            st.markdown("---")
            st.subheader("📋 Detailed Sales Records")

            display_df = filtered.copy()

            display_df["billing_date"] = (
                display_df["billing_date"]
                .dt.strftime("%Y-%m-%d")
            )

            display_df = display_df.rename(
                columns={
                    "billing_date": "Billing Date",
                    "sap_invoice_number": "SAP Invoice Number",
                    "sales_group": "Sales Group",
                    "customer_code": "Customer",
                    "sold_to_party": "Sold-to-Party",
                    "material_group": "Material Group",
                    "material_code": "Material Code",
                    "material_description": "Material Description",
                    "invoiced_quantity": "Invoiced Quantity",
                    "unit": "Unit",
                    "net_weight": "Net Weight",
                    "unit_of_weight": "Unit of Weight",
                    "net_value_usd": "Sales Value (LKR)",
                    "po_number": "PO Number",
                    "special_instructions": "Special Instructions",
                }
            )

            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True
            )

            export_sheets = {
                "Summary_By_Channel": channel_summary,
                "Summary_By_Item": item_summary,
                "Daily_Trend": daily,
                "Monthly_Trend": monthly,
                "Detailed_Sales": display_df,
            }

            st.download_button(
                "📥 Download Complete Sales Report (Excel)",
                data=multi_sheet_excel(export_sheets),
                file_name="Modern_Trade_Sales_Analysis.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )



# =========================================================
# ARPICO PO PDF UPLOAD
# =========================================================

elif main_menu == "📄 Arpico PO PDF Upload (Admin)":

    import io
    import pdfplumber

    st.header("📄 Arpico Purchase Order PDF Upload")

    if check_admin_password():

        uploaded_pdf = st.file_uploader(
            "Upload Arpico Purchase Order PDF",
            type=["pdf"],
            key="arpico_po_pdf"
        )

        if uploaded_pdf is not None:

            try:
                with pdfplumber.open(
                    io.BytesIO(uploaded_pdf.getvalue())
                ) as pdf:
                    page_texts = [
                        page.extract_text() or ""
                        for page in pdf.pages
                    ]

                all_text = "\n".join(page_texts)

                if not all_text.strip():
                    st.error(
                        "No selectable text found. "
                        "This PDF may require OCR."
                    )
                else:
                   
                    date_match = re.search(
                        r"\bDate\s+(\d{2}/\d{2}/\d{4})",
                        all_text,
                        re.IGNORECASE
                    )

                    # The displayed PDF text contains a supplier
                    # code, not a clearly labelled PO number.
                    po_match = re.search(
                        r"\b(?:PO\s*(?:No\.?|Number)|Purchase Order No\.?)"
                        r"\s*[:#-]?\s*(\d+)",
                        all_text,
                        re.IGNORECASE
                    )

                    supplier_match = re.search(
                        r"(?m)^\s*Supplier\s+\d+\s+(.+?)\s*(?:“|$)",
                        all_text
                    )

                    po_date = None

                    if date_match:
                        po_date = datetime.datetime.strptime(
                            date_match.group(1),
                            "%d/%m/%Y"
                        ).date()

                    supplier = (
                        supplier_match.group(1).strip()
                        if supplier_match else None
                    )

                    # Use the PDF filename as a fallback identifier
                    # because a separate PO number isn't visible
                    # in the extracted text.
                    po_number = (
                        po_match.group(1)
                        if po_match
                        else (
                            f"ARPICO-{po_date.isoformat()}-"
                            f"{uploaded_pdf.name.rsplit('.', 1)[0]}"
                            if po_date else None
                        )
                    )

                    # -------------------------------------
                    # 2. PARSE ITEM ROWS
                    # -------------------------------------

             
                    item_pattern = re.compile(
                        r"^\s*"
                        r"(?P<order_no>\d+)\s+"
                        r"(?P<plu>\d+)\s+"
                        r"(?P<item_code>GRRG\S+)\s+"
                        r"(?P<description>.*?)\s+"
                        r"(?P<unitize>\d+)\s+"
                        r"(?P<remark>.*?)\s+"
                        r"(?P<rate>[\d,]+\.\d{3})\s+"
                        r"(?P<discount>[\d,]+\.\d{2})\s+"
                        r"(?P<vat>[\d,]+\.\d{2})\s+"
                        r"(?P<qty>[\d,]+(?:\.\d+)?)\s+"
                        r"(?P<amount>[\d,]+\.\d{3})"
                        r"\s*$"
                    )

                    rows = []
                    current_outlet = None
                    current_sub_outlet = None
                    parse_warnings = []

                    for line in all_text.splitlines():

                        line = line.strip()

                        # Detect outlet headings such as:
                        # 02 Supply Kadawatha SC
                        
                        outlet_match = re.match(
                            r"^Supply\s+\d{1,3}\s+(.+?)\s*$",
                            line,
                            re.IGNORECASE
                        )


                        if outlet_match:
                            current_outlet = (
                                outlet_match.group(1).strip()
                            )
                            current_sub_outlet = None
                            continue

                        # Central Warehouse has separate
                        # Auto Ordering-location headings.
                        if (
                            current_outlet
                            and current_outlet.lower()
                            == "central warehouse"
                            and line.lower().startswith("auto ordering-")
                        ):
                            current_sub_outlet = line.split(
                                "-", 1
                            )[1].strip()
                            continue

                        # Ignore ordinary headings and totals.
                   
                        if not re.match(r"^\d+\s+\d+\s+GRRG", line):
                            continue



                        item_match = item_pattern.match(line)

                        if not item_match:
                            parse_warnings.append(line)
                            continue

                        d = item_match.groupdict()

                        remark = (d.get("remark") or "").strip()

                        # For Central Warehouse, use the
                        # individual destination as the outlet.
                        if (
                            current_outlet
                            and current_outlet.lower()
                            == "central warehouse"
                        ):
                            outlet_name = (
                                remark
                                or current_sub_outlet
                                or "Central Warehouse"
                            )
                        else:
                            outlet_name = current_outlet

                        if not outlet_name:
                            parse_warnings.append(
                                "Outlet not detected: " + line
                            )
                            continue

                        def to_number(value):
                            return float(value.replace(",", ""))

                        rows.append({
                            "po_number": po_number,
                            "po_date": (
                                po_date.isoformat()
                                if po_date else None
                            ),
                            "supplier": supplier,
                            "outlet_name": outlet_name,
                            "order_no": d["order_no"],
                            "item_code": d["item_code"],
                            "item_name": d["description"].strip(),
                            "rate": to_number(d["rate"]),
                            "ordered_qty": to_number(d["qty"]),
                            "amount": to_number(d["amount"]),
                            "vat": to_number(d["vat"]),
                            "discount": to_number(d["discount"]),
                            "plu": d["plu"],
                            "unitize": to_number(d["unitize"]),
                            "remark": remark or None,
                            "source_file": uploaded_pdf.name
                        })

                    # -------------------------------------
                    # 3. PREVIEW BEFORE SAVING
                    # -------------------------------------

                    st.subheader("Extracted PO Items")

                    if not rows:
                        st.error(
                            "No item rows were extracted. "
                            "Check the PDF text and item layout."
                        )

                    else:
                        preview_df = pd.DataFrame(rows)

                        st.success(
                            f"Extracted {len(preview_df)} item rows "
                            f"from {preview_df['outlet_name'].nunique()} "
                            f"outlets/destinations."
                        )

                        st.dataframe(
                            preview_df,
                            use_container_width=True,
                            hide_index=True
                        )

                        if parse_warnings:
                            with st.expander(
                                f"Rows requiring review "
                                f"({len(parse_warnings)})"
                            ):
                                st.warning(
                                    "These item-like lines were not "
                                    "parsed. Review them before saving."
                                )
                                st.code(
                                    "\n".join(parse_warnings[:100])
                                )

                        st.download_button(
                            "Download Preview as CSV",
                            data=preview_df.to_csv(
                                index=False
                            ).encode("utf-8-sig"),
                            file_name=(
                                f"arpico_po_{po_number or 'preview'}.csv"
                            ),
                            mime="text/csv"
                        )

                        # Do not save incomplete headers or
                        # rows that failed to parse.
                        can_save = bool(
                            po_number
                            and po_date
                            and not parse_warnings
                        )

                        if not can_save:
                            st.warning(
                                "Saving is disabled because the PO "
                                "header is incomplete or some item "
                                "rows need review."
                            )

                        if st.button(
                            "💾 Save PO Items to Supabase",
                            type="primary",
                            disabled=not can_save,
                            key="save_arpico_po"
                        ):
                            try:
                                response = (
                                    supabase
                                    .table("arpico_po_items")
                                    .upsert(
                                        rows,
                                        on_conflict=(
                                            "po_number,order_no,item_code"
                                        )
                                    )
                                    .execute()
                                )

                                st.success(
                                    f"PO {po_number}: "
                                    f"{len(rows)} item rows saved "
                                    "successfully."
                                )

                            except Exception as save_error:
                                st.error(
                                    "Could not save PO items: "
                                    f"{save_error}"
                                )
                                st.caption(
                                    "Check that the arpico_po_items "
                                    "table exists and that the "
                                    "Supabase client has permission "
                                    "to insert/update rows."
                                )

                    with st.expander("View extracted PDF text"):
                        st.text_area(
                            "PDF text",
                            all_text[:20000],
                            height=300
                        )

            except ImportError:
                st.error(
                    "pdfplumber is missing. Add pdfplumber "
                    "to requirements.txt."
                )

            except Exception as e:
                st.error(f"Could not process the PDF: {e}")
                
# =========================================================
# ARPICO PO ANALYSIS
# =========================================================

elif main_menu == "📊 Arpico PO Analysis":

    import pandas as pd
    import streamlit as st

    st.header("📊 Arpico PO Analysis")

    if check_admin_password():

        try:
            response = (
                supabase
                .table("arpico_po_items")
                .select("*")
                .order("po_date", desc=True)
                .execute()
            )

            data = response.data or []

            if not data:
                st.info("No Arpico PO records found.")
            else:
                df = pd.DataFrame(data)

                # Convert numeric fields safely
                for col in ["amount", "ordered_qty", "rate"]:
                    if col in df.columns:
                        df[col] = pd.to_numeric(
                            df[col], errors="coerce"
                        ).fillna(0)

                if "po_date" in df.columns:
                    df["po_date"] = pd.to_datetime(
                        df["po_date"], errors="coerce"
                    )

              
                # Filters
                st.subheader("Filters")

                c1, c2, c3 = st.columns(3)

                with c1:
                    outlets = sorted(
                        df["outlet_name"]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    )

                    selected_outlets = st.multiselect(
                        "Select outlets",
                        outlets,
                        default=outlets
                    )

                with c2:
                    po_options = sorted(
                        df["po_number"]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist(),
                        reverse=True
                    )

                    selected_pos = st.multiselect(
                        "Select PO number",
                        po_options,
                        default=po_options
                    )

                with c3:
                    items = sorted(
                        df["item_name"]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    )

                    selected_items = st.multiselect(
                        "Select items",
                        items,
                        default=items
                    )

                filtered = df[
                    df["outlet_name"].astype(str).isin(
                        selected_outlets
                    )
                    & df["po_number"].astype(str).isin(
                        selected_pos
                    )
                    & df["item_name"].astype(str).isin(
                        selected_items
                    )
                ].copy()

                # Summary
                total_value = filtered["amount"].sum()
                outlet_count = filtered["outlet_name"].nunique()
                item_count = filtered["item_code"].nunique()
                total_qty = filtered["ordered_qty"].sum()

                m1, m2, m3, m4 = st.columns(4)

                m1.metric("Total PO Value", f"Rs. {total_value:,.2f}")
                m2.metric("Outlets", f"{outlet_count:,}")
                m3.metric("Items", f"{item_count:,}")
                m4.metric("Ordered Qty", f"{total_qty:,.0f}")

                # Outlet-wise report
                st.subheader("1. Outlet-wise PO Values")

                outlet_summary = (
                    filtered.groupby("outlet_name", as_index=False)
                    .agg(
                        PO_Value=("amount", "sum"),
                        Ordered_Qty=("ordered_qty", "sum"),
                        Item_Count=("item_code", "nunique"),
                        PO_Count=("po_number", "nunique")
                    )
                    .sort_values("PO_Value", ascending=False)
                )

                st.dataframe(
                    outlet_summary,
                    use_container_width=True,
                    hide_index=True
                )

                # Item-wise report
                st.subheader("2. Outlet-wise Item Analysis")

                if not filtered.empty:
                    selected_outlet = st.selectbox(
                        "Select an outlet",
                        sorted(
                            filtered["outlet_name"]
                            .dropna()
                            .astype(str)
                            .unique()
                            .tolist()
                        )
                    )

                    outlet_items = filtered[
                        filtered["outlet_name"].astype(str)
                        == selected_outlet
                    ]

                    item_summary = (
                        outlet_items.groupby(
                            ["item_code", "item_name"],
                            as_index=False
                        )
                        .agg(
                            Ordered_Qty=("ordered_qty", "sum"),
                            Rate=("rate", "mean"),
                            PO_Value=("amount", "sum")
                        )
                        .sort_values("PO_Value", ascending=False)
                    )

                    st.dataframe(
                        item_summary,
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("No records match the selected filters.")

                # Excel download
                st.subheader("3. Download Report")

                output = io.BytesIO()

                with pd.ExcelWriter(
                    output, engine="openpyxl"
                ) as writer:
                    outlet_summary.to_excel(
                        writer,
                        sheet_name="Outlet Summary",
                        index=False
                    )

                    filtered.to_excel(
                        writer,
                        sheet_name="PO Details",
                        index=False
                    )

                st.download_button(
                    "⬇️ Download Excel Report",
                    data=output.getvalue(),
                    file_name="Arpico_PO_Analysis.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument"
                        ".spreadsheetml.sheet"
                    )
                )

        except Exception as e:
            st.error(f"Could not load Arpico PO data: {e}")


elif main_menu == "📤 Upload New Stock (Admin)":

    st.header("📤 Upload Daily Stock File")

    if check_admin_password():

        col1, col2 = st.columns(2)

        with col1:
            upload_date = st.date_input(
                "Stock Data Date",
                datetime.date.today()
            )

        with col2:
            batch_num = st.selectbox(
                "Upload Batch",
                [
                    "Batch 1",
                    "Batch 2",
                    "Batch 3",
                    "Batch 4",
                    "Batch 5",
                ]
            )

        note = st.text_input("Optional Note")

        badge = f"{upload_date.isoformat()} - {batch_num}"

        if note.strip():
            badge += f" ({note.strip()})"

        st.info(f"Upload batch: {badge}")

        if badge in stock_badges:
            st.warning(
                "This batch already exists. Uploading will replace it."
            )

        stock_file = st.file_uploader(
            "Choose Stock Excel File",
            type=["xlsx", "xls"],
            key="stock_upload"
        )

        if stock_file and st.button("Save Stock to Database"):

            try:
                df = pd.read_excel(stock_file)

                rename_dict = {
                    "SKU Description": "SKU_Description",
                    "Store Description": "Store_Description",
                    "Current Stock On Hand Units": "Current_Stock_Units",
                    "Material Status Description": "Material_Status_Desc",
                    "Last Update Date Time": "Last_Update_Time",
                }

                df = df.rename(columns=rename_dict)

                if "SKU" not in df.columns or "Store" not in df.columns:
                    st.error(
                        "Excel file must contain SKU and Store columns."
                    )
                    st.stop()

                df["SKU"] = df["SKU"].astype(str).apply(
                    lambda x: x.split(".")[0].strip()
                )

                df["Category"] = df["SKU"].apply(categorize_by_sku)

                if "Current_Stock_Units" in df.columns:
                    df["Current_Stock_Units"] = pd.to_numeric(
                        df["Current_Stock_Units"],
                        errors="coerce"
                    ).fillna(0)

                df["Uploaded_At"] = badge

                valid_columns = [
                    "Uploaded_At",
                    "Store",
                    "Store_Description",
                    "SKU",
                    "SKU_Description",
                    "Category",
                    "Current_Stock_Units",
                    "Material_Status_Desc",
                    "Last_Update_Time",
                ]

                df = df[
                    [c for c in valid_columns if c in df.columns]
                ]

                df = df.astype(object).where(
                    pd.notnull(df),
                    None
                )

                if badge in stock_badges:
                    supabase.table("stock_history").delete().eq(
                        "Uploaded_At",
                        badge
                    ).execute()

                records = df.to_dict(orient="records")
                progress = st.progress(0)

                for i in range(0, len(records), 200):
                    chunk = records[i:i + 200]

                    safe_chunk = [
                        {
                            key: make_json_safe(value)
                            for key, value in record.items()
                        }
                        for record in chunk
                    ]

                    supabase.table("stock_history").insert(
                        safe_chunk
                    ).execute()

                    progress.progress(
                        min((i + len(chunk)) / len(records), 1.0)
                    )

                st.cache_data.clear()
                st.success(f"Stock upload completed: {badge}")
                st.rerun()

            except Exception as e:
                st.error(f"Stock upload failed: {e}")


# =========================================================
# 4. MANAGE / DELETE STOCK UPLOADS
# =========================================================

elif main_menu == "🗑️ Manage / Delete Uploads (Admin)":

    st.header("🗑️ Manage Stock Uploads")

    if check_admin_password():

        if not stock_badges:
            st.info("No stock uploads found.")

        else:
            selected_batch = st.selectbox(
                "Select stock batch to delete",
                stock_badges
            )

            st.warning(
                f"Selected batch '{selected_batch}' will be permanently deleted."
            )

            if st.button("Delete Selected Stock Batch"):

                try:
                    supabase.table("stock_history").delete().eq(
                        "Uploaded_At",
                        selected_batch
                    ).execute()

                    st.cache_data.clear()
                    st.success("Stock batch deleted.")
                    st.rerun()

                except Exception as e:
                    st.error(f"Delete failed: {e}")


# =========================================================
# 5. MANAGE / DELETE SALES UPLOADS
# =========================================================

elif main_menu == "🗑️ Manage Sales Uploads (Admin)":

    st.header("🗑️ Manage Sales Uploads")

    if check_admin_password():

        try:
            sales_df = fetch_all_rows("sales_history")

            if sales_df.empty:
                st.info("No sales uploads found.")

            else:
                batches = sorted(
                    sales_df["upload_batch"]
                    .dropna()
                    .unique()
                    .tolist(),
                    reverse=True
                )

                batch_summary = (
                    sales_df.groupby("upload_batch")
                    .agg(
                        Rows=("row_hash", "count"),
                        From_Date=("billing_date", "min"),
                        To_Date=("billing_date", "max"),
                    )
                    .reset_index()
                )

                st.dataframe(
                    batch_summary,
                    use_container_width=True,
                    hide_index=True
                )

                selected_sales_batch = st.selectbox(
                    "Select sales batch to delete",
                    batches
                )

                st.warning(
                    "This permanently deletes every sales row in the "
                    "selected upload batch."
                )

                if st.button("Delete Selected Sales Batch"):

                    supabase.table("sales_history").delete().eq(
                        "upload_batch",
                        selected_sales_batch
                    ).execute()

                    st.cache_data.clear()
                    st.success("Sales batch deleted.")
                    st.rerun()

        except Exception as e:
            st.error(f"Could not load sales batches: {e}")


# =========================================================
# 6. HISTORICAL OOS TREND ANALYSIS
# =========================================================

elif main_menu == "📈 Historical OOS Trend Analysis":

    st.header("📈 Historical OOS Trend Analysis")

    if not stock_badges:
        st.warning("No stock history found. Upload stock files first.")

    else:
        col1, col2 = st.columns(2)

        with col1:
            start_date = st.date_input(
                "Start Date",
                datetime.date.today() - datetime.timedelta(days=7),
                key="oos_start"
            )

        with col2:
            end_date = st.date_input(
                "End Date",
                datetime.date.today(),
                key="oos_end"
            )

        if start_date > end_date:
            st.error("Start Date cannot be after End Date.")
            st.stop()

        filtered_badges = []

        for badge in stock_badges:
            try:
                badge_date = datetime.datetime.strptime(
                    badge.split(" - ")[0],
                    "%Y-%m-%d"
                ).date()

                if start_date <= badge_date <= end_date:
                    filtered_badges.append(badge)

            except Exception:
                continue

        if not filtered_badges:
            st.warning("No stock uploads found in this date range.")

        else:
            mode = st.radio(
                "Analysis Type",
                [
                    "Item-wise Analysis",
                    "Outlet-wise Analysis",
                ],
                horizontal=True
            )

            category = st.radio(
                "Category",
                ["Dairies", "Rice"],
                horizontal=True,
                key="oos_category"
            )

            sample = fetch_all_batch_data(filtered_badges[-1])

            if sample.empty:
                st.warning("No stock records found.")
                st.stop()

            store_col = (
                "Store_Description"
                if "Store_Description" in sample.columns
                else "Store"
            )

            item_col = (
                "SKU_Description"
                if "SKU_Description" in sample.columns
                else "SKU"
            )

            if mode == "Item-wise Analysis":

                cat_sample = sample[
                    sample["Category"] == category
                ]

                item_options = sorted(
                    cat_sample[item_col]
                    .dropna()
                    .astype(str)
                    .unique()
                )

                selected_item = st.selectbox(
                    "Select Item",
                    ["All Items"] + item_options
                )

            else:
                non_outlet = (
                    sample[store_col].astype(str).str.contains(
                        "DCW1|Kerawalapitiya|Common Storage|Import DC",
                        case=False,
                        na=False
                    )
                    |
                    sample["Store"].astype(str).str.contains(
                        "DCW1|IZCS|DCDI",
                        case=False,
                        na=False
                    )
                )

                outlet_options = sorted(
                    sample.loc[~non_outlet, store_col]
                    .dropna()
                    .astype(str)
                    .unique()
                )

                selected_outlet = st.selectbox(
                    "Select Outlet",
                    ["All Outlets"] + outlet_options
                )

            if st.button("Generate OOS Trend"):

                trend_rows = []

                for badge in filtered_badges:

                    batch_df = fetch_all_batch_data(badge)

                    if batch_df.empty:
                        continue

                    batch_store_col = (
                        "Store_Description"
                        if "Store_Description" in batch_df.columns
                        else "Store"
                    )

                    batch_item_col = (
                        "SKU_Description"
                        if "SKU_Description" in batch_df.columns
                        else "SKU"
                    )

                    non_outlet = (
                        batch_df[batch_store_col].astype(str).str.contains(
                            "DCW1|Kerawalapitiya|Common Storage|Import DC",
                            case=False,
                            na=False
                        )
                        |
                        batch_df["Store"].astype(str).str.contains(
                            "DCW1|IZCS|DCDI",
                            case=False,
                            na=False
                        )
                    )

                    outlet_df = batch_df[~non_outlet].copy()

                    outlet_df["Current_Stock_Units"] = pd.to_numeric(
                        outlet_df.get(
                            "Current_Stock_Units",
                            pd.Series(0, index=outlet_df.index)
                        ),
                        errors="coerce"
                    ).fillna(0)

                    zero_df = outlet_df[
                        (outlet_df["Category"] == category)
                        & (outlet_df["Current_Stock_Units"] <= 0)
                    ]

                    if mode == "Item-wise Analysis":

                        if selected_item != "All Items":
                            zero_df = zero_df[
                                zero_df[batch_item_col].astype(str)
                                == selected_item
                            ]

                        count = zero_df[batch_store_col].nunique()

                        details = ", ".join(
                            sorted(
                                zero_df[batch_store_col]
                                .dropna()
                                .astype(str)
                                .unique()
                            )
                        )

                        trend_rows.append({
                            "Batch / Date": badge,
                            "OOS Outlets Count": count,
                            "OOS Outlets": details,
                        })

                    else:

                        if selected_outlet != "All Outlets":
                            zero_df = zero_df[
                                zero_df[batch_store_col].astype(str)
                                == selected_outlet
                            ]

                        count = zero_df[batch_item_col].nunique()

                        details = ", ".join(
                            sorted(
                                zero_df[batch_item_col]
                                .dropna()
                                .astype(str)
                                .unique()
                            )
                        )

                        trend_rows.append({
                            "Batch / Date": badge,
                            "OOS Items Count": count,
                            "OOS Items": details,
                        })

                trend_df = pd.DataFrame(trend_rows)

                if trend_df.empty:
                    st.warning("No OOS data found.")

                else:
                    y_col = (
                        "OOS Outlets Count"
                        if mode == "Item-wise Analysis"
                        else "OOS Items Count"
                    )

                    fig = px.line(
                        trend_df,
                        x="Batch / Date",
                        y=y_col,
                        markers=True,
                        title=f"{category} — {mode}"
                    )

                    fig.update_layout(xaxis_tickangle=-45)

                    st.plotly_chart(
                        fig,
                        use_container_width=True
                    )

                    st.dataframe(
                        trend_df,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.download_button(
                        "📥 Download OOS Trend Excel",
                        data=dataframe_to_excel(
                            trend_df,
                            "OOS_Trend"
                        ),
                        file_name="OOS_Trend_Report.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        )
                    )


# =========================================================
# 7. LATEST STOCK DASHBOARDS
# =========================================================

elif main_menu in [
    "🔍 Outlet Stock Search (Latest)",
    "⚠️ Zero Stock Report (Latest)",
    "🏬 Warehouse Stock (Latest)",
]:

    st.header(main_menu)

    if not latest_badge:
        st.warning(
            "No stock uploads found. Upload a stock file first."
        )
        st.stop()

    st.caption(f"Latest Stock Update: {latest_badge}")

    try:
        df = fetch_all_batch_data(latest_badge)

    except Exception as e:
        st.error(f"Could not retrieve stock data: {e}")
        st.stop()

    if df.empty:
        st.warning("No stock records found in the latest upload.")
        st.stop()

    if "SKU" in df.columns:
        df["SKU"] = df["SKU"].astype(str).apply(
            lambda x: x.split(".")[0].strip()
        )

        df["Category"] = df["SKU"].apply(
            categorize_by_sku
        )

    df["Current_Stock_Units"] = pd.to_numeric(
        df.get(
            "Current_Stock_Units",
            pd.Series(0, index=df.index)
        ),
        errors="coerce"
    ).fillna(0)

    store_col = (
        "Store_Description"
        if "Store_Description" in df.columns
        else "Store"
    )

    item_col = (
        "SKU_Description"
        if "SKU_Description" in df.columns
        else "SKU"
    )

    warehouse_mask = (
        df[store_col].astype(str).str.strip().str.upper().eq("DCW1")
        |
        df["Store"].astype(str).str.strip().str.upper().eq("DCW1")
    )

    warehouse_df = df[warehouse_mask].copy()

    non_outlet_mask = (
        df[store_col].astype(str).str.contains(
            "DCW1|Kerawalapitiya|Common Storage|Import DC",
            case=False,
            na=False
        )
        |
        df["Store"].astype(str).str.contains(
            "DCW1|IZCS|DCDI",
            case=False,
            na=False
        )
    )

    outlets_df = df[~non_outlet_mask].copy()

    # OUTLET STOCK SEARCH
    if main_menu == "🔍 Outlet Stock Search (Latest)":

        outlets = sorted(
            outlets_df[store_col]
            .dropna()
            .astype(str)
            .unique()
        )

        if not outlets:
            st.warning("No outlet stock records found.")

        else:
            selected_outlet = st.selectbox(
                "Select Outlet",
                outlets
            )

            outlet_data = outlets_df[
                outlets_df[store_col].astype(str) == selected_outlet
            ]

            item_options = sorted(
                outlet_data[item_col]
                .dropna()
                .astype(str)
                .unique()
            )

            if item_options:

                selected_item = st.selectbox(
                    "Select Item",
                    item_options
                )

                details = outlet_data[
                    outlet_data[item_col].astype(str) == selected_item
                ].iloc[0]

                st.subheader(selected_item)

                c1, c2 = st.columns(2)

                c1.metric(
                    "Current Stock On Hand",
                    f"{details.get('Current_Stock_Units', 0):,.0f}"
                )

                c2.write(
                    f"**SKU:** {details.get('SKU', 'N/A')}"
                )

                c2.write(
                    f"**Category:** {details.get('Category', 'N/A')}"
                )

                st.write(
                    f"**Last System Update:** "
                    f"{details.get('Last_Update_Time', 'N/A')}"
                )

                st.write(
                    f"**Status:** "
                    f"{details.get('Material_Status_Desc', 'N/A')}"
                )

    # ZERO STOCK REPORT
    elif main_menu == "⚠️ Zero Stock Report (Latest)":

        st.subheader("Zero Stock Outlets Report")

        tab_dairy, tab_rice = st.tabs(
            ["🥛 Dairies", "🍚 Rice"]
        )

        def render_zero_stock(category):

            category_df = outlets_df[
                outlets_df["Category"] == category
            ]

            zero_df = category_df[
                category_df["Current_Stock_Units"] <= 0
            ].copy()

            if zero_df.empty:
                st.success(
                    f"No zero-stock {category} items found."
                )
                return

            item_options = [
                "All Zero Stock Items"
            ] + sorted(
                zero_df[item_col]
                .dropna()
                .astype(str)
                .unique()
            )

            selected_zero_item = st.selectbox(
                f"Filter {category} Item",
                item_options,
                key=f"zero_item_{category}"
            )

            if selected_zero_item != "All Zero Stock Items":
                zero_df = zero_df[
                    zero_df[item_col].astype(str)
                    == selected_zero_item
                ]

            st.metric(
                "Zero Stock Records",
                f"{len(zero_df):,}"
            )

            display_cols = [
                store_col,
                "SKU",
                item_col,
                "Current_Stock_Units",
                "Material_Status_Desc",
            ]

            display_cols = [
                col for col in display_cols
                if col in zero_df.columns
            ]

            report = zero_df[display_cols].copy()

            st.dataframe(
                report,
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                f"Download {category} Zero Stock Excel",
                data=dataframe_to_excel(
                    report,
                    category
                ),
                file_name=f"{category}_Zero_Stock.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key=f"download_zero_{category}"
            )

        with tab_dairy:
            render_zero_stock("Dairies")

        with tab_rice:
            render_zero_stock("Rice")

    # WAREHOUSE STOCK
    elif main_menu == "🏬 Warehouse Stock (Latest)":

        st.subheader("Warehouse Stock — DCW1")

        if warehouse_df.empty:
            st.warning(
                "No DCW1 warehouse records found."
            )

        else:
            warehouse_df = warehouse_df.drop_duplicates(
                subset=["SKU"],
                keep="last"
            )

            display_cols = [
                "SKU",
                item_col,
                "Current_Stock_Units",
                "Category",
            ]

            display_cols = [
                col for col in display_cols
                if col in warehouse_df.columns
            ]

            warehouse_display = warehouse_df[display_cols].copy()

            st.dataframe(
                warehouse_display,
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                "📥 Download Warehouse Stock Excel",
                data=dataframe_to_excel(
                    warehouse_display,
                    "Warehouse_Stock"
                ),
                file_name="Warehouse_Stock.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )
