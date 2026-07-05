import os
from datetime import datetime
from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Spare Parts Customer Portal", page_icon="⚙️", layout="wide")

DATA_DIR = "data"
DEFAULT_INVENTORY = os.path.join(DATA_DIR, "inventory.xlsx")
RFQ_FILE = os.path.join(DATA_DIR, "customer_rfqs.xlsx")
ADMIN_PASSWORD = "Manutd@033"  # Change this before deployment

SAFE_COLUMNS = ["Part Number", "Description", "Manufacturer", "Availability"]

COLUMN_CANDIDATES = {
    "part": ["Part Number", "Name", "Internal Reference", "Item", "Item Code", "Part No", "Part No.", "P/N"],
    "description": ["Description", "Desc", "Product Description", "Item Description", "DESCRIPTION"],
    "manufacturer": ["Manufacturer", "OEM", "Brand", "Make", "Vendor"],
    "category": ["Product Category", "Category", "Type"],
    "qty": ["Quantity On Hand", "Qty", "QTY", "Quantity", "Stock", "On Hand"],
    "cost": ["Cost", "Unit Cost", "Average Cost"],
    "total": ["Total Cost", "Total Value", "Value"],
}


def first_existing(cols, candidates):
    normalized = {str(c).strip().lower(): c for c in cols}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in normalized:
            return normalized[key]
    return None


@st.cache_data(show_spinner=False)
def load_inventory(path: str):
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]

    part_col = first_existing(df.columns, COLUMN_CANDIDATES["part"])
    desc_col = first_existing(df.columns, COLUMN_CANDIDATES["description"])
    man_col = first_existing(df.columns, COLUMN_CANDIDATES["manufacturer"])
    cat_col = first_existing(df.columns, COLUMN_CANDIDATES["category"])
    qty_col = first_existing(df.columns, COLUMN_CANDIDATES["qty"])
    cost_col = first_existing(df.columns, COLUMN_CANDIDATES["cost"])
    total_col = first_existing(df.columns, COLUMN_CANDIDATES["total"])

    out = pd.DataFrame()
    out["Part Number"] = df[part_col].astype(str).str.strip() if part_col else ""
    out["Description"] = df[desc_col].fillna("TBA").astype(str).str.strip() if desc_col else "TBA"
    out["Manufacturer"] = df[man_col].fillna("TBA").astype(str).str.strip() if man_col else "TBA"
    out["Category"] = df[cat_col].fillna("TBA").astype(str).str.strip() if cat_col else "TBA"
    out["Qty"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0) if qty_col else 0
    out["Cost"] = pd.to_numeric(df[cost_col], errors="coerce").fillna(0) if cost_col else 0
    out["Total Value"] = pd.to_numeric(df[total_col], errors="coerce").fillna(out["Qty"] * out["Cost"]) if total_col else out["Qty"] * out["Cost"]
    out = out[out["Part Number"].astype(str).str.lower().ne("nan")]
    out = out[out["Part Number"].astype(str).str.strip().ne("")]
    out["Availability"] = out["Qty"].apply(lambda q: "Limited Stock" if q <= 2 and q > 0 else ("Available" if q > 2 else "Out of Stock"))
    return out


def save_uploaded_inventory(uploaded_file):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(DEFAULT_INVENTORY, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.cache_data.clear()


def load_rfqs():
    if os.path.exists(RFQ_FILE):
        return pd.read_excel(RFQ_FILE)
    return pd.DataFrame(columns=["Date", "Company", "Contact", "Email", "Mobile", "Part Number", "Required Qty", "Notes", "Status"])


def save_rfq(row):
    os.makedirs(DATA_DIR, exist_ok=True)
    rfqs = load_rfqs()
    rfqs = pd.concat([rfqs, pd.DataFrame([row])], ignore_index=True)
    rfqs.to_excel(RFQ_FILE, index=False)


def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]
        for i, col in enumerate(df.columns):
            width = min(max(len(str(col)) + 4, int(df[col].astype(str).str.len().quantile(0.9)) + 2 if len(df) else 14), 45)
            ws.set_column(i, i, width)
    output.seek(0)
    return output.getvalue()


def search_df(df, query, manufacturer, availability):
    f = df.copy()
    if query:
        q = query.lower().strip()
        mask = (
            f["Part Number"].astype(str).str.lower().str.contains(q, na=False) |
            f["Description"].astype(str).str.lower().str.contains(q, na=False) |
            f["Manufacturer"].astype(str).str.lower().str.contains(q, na=False) |
            f["Category"].astype(str).str.lower().str.contains(q, na=False)
        )
        f = f[mask]
    if manufacturer and manufacturer != "All":
        f = f[f["Manufacturer"] == manufacturer]
    if availability and availability != "All":
        f = f[f["Availability"] == availability]
    return f


def header():
    st.title("⚙️ Industrial Spare Parts Customer Portal")
    st.caption("Ready-stock spare parts inquiry portal | Saudi Arabia, GCC & Middle East")


if not os.path.exists(DEFAULT_INVENTORY):
    st.error("The inventory.xlsx file was not found inside the data folder.")
    st.stop()

inventory = load_inventory(DEFAULT_INVENTORY)

page = st.sidebar.radio("Menu", ["Customer Portal", "Request Quotation", "Admin Dashboard", "Upload Inventory", "RFQ Inbox"])

if page == "Customer Portal":
    header()
    st.info("Search by part number, manufacturer, or description. Prices are not displayed here; please submit an RFQ and we will respond to you promptly.")
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        query = st.text_input("Search Part Number / Description / Manufacturer")
    with c2:
        mans = ["All"] + sorted([x for x in inventory["Manufacturer"].dropna().unique().tolist() if str(x).strip()])
        manufacturer = st.selectbox("Manufacturer", mans)
    with c3:
        availability = st.selectbox("Availability", ["All", "Available", "Limited Stock"])

    visible = inventory[inventory["Qty"] > 0]
    result = search_df(visible, query, manufacturer, availability)
    st.metric("Matching Available Items", len(result))
    st.dataframe(result[SAFE_COLUMNS], use_container_width=True, hide_index=True)

    st.download_button("Download Customer List (No Prices)", to_excel_bytes(result[SAFE_COLUMNS]), "customer_stock_list.xlsx")

elif page == "Request Quotation":
    header()
    st.subheader("Request for Quotation")
    with st.form("rfq_form"):
        company = st.text_input("Company Name *")
        contact = st.text_input("Contact Person *")
        email = st.text_input("Email *")
        mobile = st.text_input("Mobile / WhatsApp")
        part = st.text_input("Part Number *")
        qty = st.number_input("Required Qty", min_value=1, value=1)
        notes = st.text_area("Notes")
        submitted = st.form_submit_button("Submit RFQ")
        if submitted:
            if not company or not contact or not email or not part:
                st.error("Please complete all required fields marked with *.")
            else:
                save_rfq({
                    "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Company": company,
                    "Contact": contact,
                    "Email": email,
                    "Mobile": mobile,
                    "Part Number": part,
                    "Required Qty": qty,
                    "Notes": notes,
                    "Status": "New",
                })
                st.success("Your request has been submitted successfully. We will contact you shortly.")

elif page == "Admin Dashboard":
    pwd = st.sidebar.text_input("Admin Password", type="password")
    if pwd != ADMIN_PASSWORD:
        st.warning("Please enter the admin password from the sidebar.")
        st.stop()
    header()
    st.subheader("Admin Dashboard")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Items", len(inventory))
    c2.metric("Total Qty", f"{inventory['Qty'].sum():,.0f}")
    c3.metric("Manufacturers", inventory["Manufacturer"].nunique())
    c4.metric("TBA Descriptions", int((inventory["Description"].str.upper() == "TBA").sum()))

    st.divider()
    st.subheader("Internal Inventory (Includes Cost)")
    st.dataframe(inventory, use_container_width=True, hide_index=True)
    st.download_button("Download Full Internal Inventory", to_excel_bytes(inventory), "internal_inventory.xlsx")

elif page == "Upload Inventory":
    pwd = st.sidebar.text_input("Admin Password", type="password")
    if pwd != ADMIN_PASSWORD:
        st.warning("Please enter the admin password from the sidebar.")
        st.stop()
    header()
    st.subheader("Upload New Inventory Excel")
    st.warning("The current inventory.xlsx file will be replaced with the newly uploaded file.")
    uploaded = st.file_uploader("Choose Excel file", type=["xlsx"])
    if uploaded and st.button("Replace Inventory"):
        save_uploaded_inventory(uploaded)
        st.success("The inventory file has been updated successfully. Please refresh the page if the data does not update immediately.")

elif page == "RFQ Inbox":
    pwd = st.sidebar.text_input("Admin Password", type="password")
    if pwd != ADMIN_PASSWORD:
        st.warning("Please enter the admin password from the sidebar.")
        st.stop()
    header()
    st.subheader("Customer RFQs")
    rfqs = load_rfqs()
    st.metric("RFQs", len(rfqs))
    st.dataframe(rfqs, use_container_width=True, hide_index=True)
    st.download_button("Download RFQs", to_excel_bytes(rfqs), "customer_rfqs.xlsx")
