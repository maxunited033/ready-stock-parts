import os
from datetime import datetime
from io import BytesIO
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Ready Stock Parts | Industrial OEM Spare Parts",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = "data"
DEFAULT_INVENTORY = os.path.join(DATA_DIR, "inventory.xlsx")
RFQ_FILE = os.path.join(DATA_DIR, "customer_rfqs.xlsx")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "Manutd@033")
CONTACT_EMAIL = "mossab.rozi@gmail.com"
CONTACT_MOBILE = "+966561261005"
WHATSAPP_NUMBER = "966561261005"

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


def inject_css():
    st.markdown(
        """
        <style>
        #MainMenu, footer, header {visibility: hidden;}
        .block-container {padding-top: 1.2rem; padding-bottom: 3rem;}
        .hero {
            padding: 2.2rem 2rem;
            border-radius: 22px;
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 45%, #0ea5e9 100%);
            color: white;
            margin-bottom: 1.2rem;
            box-shadow: 0 16px 45px rgba(15, 23, 42, 0.18);
        }
        .hero h1 {font-size: 2.7rem; margin-bottom: 0.4rem; font-weight: 800; letter-spacing: -0.04em;}
        .hero p {font-size: 1.08rem; opacity: 0.94; max-width: 980px;}
        .badge {
            display: inline-block; padding: 0.35rem 0.7rem; border-radius: 999px;
            background: rgba(255,255,255,0.16); border: 1px solid rgba(255,255,255,0.25);
            margin-right: 0.35rem; margin-bottom: 0.5rem; font-size: 0.85rem;
        }
        .card {
            border: 1px solid #e5e7eb; border-radius: 18px; padding: 1.2rem;
            background: white; box-shadow: 0 8px 28px rgba(15,23,42,0.06); height: 100%;
        }
        .card h3 {margin-top: 0; color: #0f172a;}
        .muted {color: #64748b;}
        .section-title {font-size: 1.4rem; font-weight: 800; color: #0f172a; margin: 1.2rem 0 0.8rem;}
        .brand-pill {
            display: inline-block; padding: 0.55rem 0.85rem; border: 1px solid #dbeafe;
            border-radius: 999px; margin: 0.22rem; background: #eff6ff; color: #1e3a8a; font-weight: 650;
        }
        .footer-box {border-top: 1px solid #e5e7eb; margin-top: 2rem; padding-top: 1.2rem; color: #64748b;}
        .small-note {font-size:0.88rem; color:#64748b;}
        .availability-available {color:#15803d; font-weight:700;}
        .availability-limited {color:#b45309; font-weight:700;}
        </style>
        """,
        unsafe_allow_html=True,
    )


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
    out["Availability"] = out["Qty"].apply(lambda q: "Limited Stock" if 0 < q <= 2 else ("Available" if q > 2 else "Out of Stock"))
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
            width = min(max(len(str(col)) + 4, int(df[col].astype(str).str.len().quantile(0.9)) + 2 if len(df) else 14), 48)
            ws.set_column(i, i, width)
    output.seek(0)
    return output.getvalue()


def search_df(df, query, manufacturer, availability):
    f = df.copy()
    if query:
        q = query.lower().strip()
        mask = (
            f["Part Number"].astype(str).str.lower().str.contains(q, na=False, regex=False) |
            f["Description"].astype(str).str.lower().str.contains(q, na=False, regex=False) |
            f["Manufacturer"].astype(str).str.lower().str.contains(q, na=False, regex=False) |
            f["Category"].astype(str).str.lower().str.contains(q, na=False, regex=False)
        )
        f = f[mask]
    if manufacturer and manufacturer != "All Manufacturers":
        f = f[f["Manufacturer"] == manufacturer]
    if availability and availability != "All Availability":
        f = f[f["Availability"] == availability]
    return f


def hero():
    st.markdown(
        """
        <div class="hero">
            <span class="badge">Saudi Arabia</span>
            <span class="badge">GCC & Middle East Supply</span>
            <span class="badge">Industrial OEM Spare Parts</span>
            <h1>Ready Stock Parts</h1>
            <p>Professional B2B portal for industrial spare parts availability checks and RFQ requests. Search ready-stock items by part number, OEM, manufacturer, or description.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def contact_box():
    whatsapp_text = quote_plus("Hello Mossab, I would like to inquire about industrial spare parts availability.")
    whatsapp_url = f"https://wa.me/{WHATSAPP_NUMBER}?text={whatsapp_text}"
    st.markdown(
        f"""
        <div class="card">
            <h3>Contact Ready Stock Parts</h3>
            <p class="muted">For urgent requirements, RFQs, or technical spare parts identification.</p>
            <p><b>Email:</b> {CONTACT_EMAIL}</p>
            <p><b>Mobile / WhatsApp:</b> {CONTACT_MOBILE}</p>
            <p><a href="{whatsapp_url}" target="_blank">Open WhatsApp Inquiry</a></p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def require_admin():
    pwd = st.sidebar.text_input("Admin Password", type="password")
    if pwd != ADMIN_PASSWORD:
        st.warning("Please enter the admin password from the sidebar to access this section.")
        st.stop()


inject_css()

if not os.path.exists(DEFAULT_INVENTORY):
    st.error("Inventory file not found. Please make sure data/inventory.xlsx exists in the project.")
    st.stop()

inventory = load_inventory(DEFAULT_INVENTORY)
public_inventory = inventory[inventory["Qty"] > 0].copy()

st.sidebar.markdown("## Ready Stock Parts")
st.sidebar.caption("Industrial OEM Spare Parts Portal")
page = st.sidebar.radio(
    "Navigation",
    ["Home", "Search Parts", "Request Quotation", "OEM Brands", "About", "Admin Dashboard", "Upload Inventory", "RFQ Inbox"],
)
st.sidebar.divider()
st.sidebar.caption(f"Email: {CONTACT_EMAIL}")
st.sidebar.caption(f"WhatsApp: {CONTACT_MOBILE}")

if page == "Home":
    hero()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Available Line Items", f"{len(public_inventory):,}")
    c2.metric("OEM / Manufacturers", f"{public_inventory['Manufacturer'].nunique():,}")
    c3.metric("Markets Served", "GCC + ME")
    c4.metric("RFQ Response", "Fast")

    st.markdown('<div class="section-title">What We Provide</div>', unsafe_allow_html=True)
    a, b, c = st.columns(3)
    with a:
        st.markdown('<div class="card"><h3>Ready-Stock Availability</h3><p class="muted">Search and request availability for industrial spare parts from existing inventory.</p></div>', unsafe_allow_html=True)
    with b:
        st.markdown('<div class="card"><h3>OEM Spare Parts Support</h3><p class="muted">Support for parts identification using part number, description, manufacturer, and equipment context.</p></div>', unsafe_allow_html=True)
    with c:
        st.markdown('<div class="card"><h3>RFQ Workflow</h3><p class="muted">Submit quotation requests and receive commercial follow-up through email or WhatsApp.</p></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Featured OEM / Manufacturer Categories</div>', unsafe_allow_html=True)
    for brand in ["Sundyne", "Milton Roy", "Turck", "Banner Engineering", "HMD Kontro", "Autronica", "3D Instruments", "Meriam", "Zenith Controls"]:
        st.markdown(f'<span class="brand-pill">{brand}</span>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Start Your Search</div>', unsafe_allow_html=True)
    st.info("Use the Search Parts page to check public availability, or submit an RFQ directly from the Request Quotation page.")

elif page == "Search Parts":
    hero()
    st.subheader("Search Available Spare Parts")
    st.caption("Cost, internal value, and sensitive inventory details are not displayed in the customer portal.")

    c1, c2, c3 = st.columns([2.2, 1, 1])
    with c1:
        query = st.text_input("Search by part number, description, OEM, manufacturer, or category")
    with c2:
        mans = ["All Manufacturers"] + sorted([x for x in public_inventory["Manufacturer"].dropna().unique().tolist() if str(x).strip()])
        manufacturer = st.selectbox("Manufacturer", mans)
    with c3:
        availability = st.selectbox("Availability", ["All Availability", "Available", "Limited Stock"])

    result = search_df(public_inventory, query, manufacturer, availability)
    st.metric("Matching Available Items", f"{len(result):,}")
    st.dataframe(result[SAFE_COLUMNS], use_container_width=True, hide_index=True)
    st.download_button("Download Customer Stock List", to_excel_bytes(result[SAFE_COLUMNS]), "ready_stock_parts_customer_list.xlsx")

    st.markdown("---")
    st.markdown("**Need pricing or delivery details?** Go to the Request Quotation page and submit the required part number and quantity.")

elif page == "Request Quotation":
    hero()
    st.subheader("Request for Quotation")
    st.caption("Submit your requirement and our team will respond with availability, pricing, and delivery options.")

    with st.form("rfq_form"):
        col1, col2 = st.columns(2)
        with col1:
            company = st.text_input("Company Name *")
            contact = st.text_input("Contact Person *")
            email = st.text_input("Email *")
            mobile = st.text_input("Mobile / WhatsApp")
        with col2:
            part = st.text_input("Part Number *")
            qty = st.number_input("Required Quantity", min_value=1, value=1)
            notes = st.text_area("Notes / Equipment Details")
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
                st.success("Your RFQ has been submitted successfully. We will contact you shortly.")

    st.markdown("---")
    contact_box()

elif page == "OEM Brands":
    hero()
    st.subheader("OEM / Manufacturer Coverage")
    st.caption("Below are manufacturers detected from the current inventory file.")
    brand_counts = public_inventory.groupby("Manufacturer", dropna=False).size().reset_index(name="Available Line Items")
    brand_counts = brand_counts.sort_values("Available Line Items", ascending=False)
    st.dataframe(brand_counts, use_container_width=True, hide_index=True)

elif page == "About":
    hero()
    left, right = st.columns([1.4, 1])
    with left:
        st.markdown(
            """
            ### About Ready Stock Parts
            Ready Stock Parts is a B2B inquiry portal for industrial OEM spare parts available from ready stock.

            We support industrial customers across Saudi Arabia, GCC countries, and the wider Middle East region with spare parts availability checks, RFQ handling, and technical identification support.

            ### Industries Served
            - Oil & Gas
            - Petrochemical
            - Water & Wastewater
            - Power Generation
            - Industrial Manufacturing
            - EPC Contractors
            - Maintenance and MRO companies
            """
        )
    with right:
        contact_box()

elif page == "Admin Dashboard":
    require_admin()
    hero()
    st.subheader("Admin Dashboard")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Inventory Items", f"{len(inventory):,}")
    c2.metric("Total Quantity", f"{inventory['Qty'].sum():,.0f}")
    c3.metric("Manufacturers", f"{inventory['Manufacturer'].nunique():,}")
    c4.metric("TBA Descriptions", f"{int((inventory['Description'].str.upper() == 'TBA').sum()):,}")

    st.divider()
    st.subheader("Internal Inventory")
    st.caption("This internal view includes cost and total value. Do not share this page with customers.")
    st.dataframe(inventory, use_container_width=True, hide_index=True)
    st.download_button("Download Full Internal Inventory", to_excel_bytes(inventory), "internal_inventory.xlsx")

elif page == "Upload Inventory":
    require_admin()
    hero()
    st.subheader("Upload New Inventory Excel")
    st.warning("Uploading a new file will replace the current data/inventory.xlsx file used by the portal.")
    uploaded = st.file_uploader("Choose Excel file", type=["xlsx"])
    if uploaded and st.button("Replace Inventory"):
        save_uploaded_inventory(uploaded)
        st.success("Inventory file updated successfully. Refresh the page if the new data is not displayed immediately.")

elif page == "RFQ Inbox":
    require_admin()
    hero()
    st.subheader("Customer RFQ Inbox")
    rfqs = load_rfqs()
    c1, c2 = st.columns(2)
    c1.metric("Total RFQs", f"{len(rfqs):,}")
    c2.metric("New RFQs", f"{(rfqs['Status'].astype(str).str.lower() == 'new').sum() if len(rfqs) else 0:,}")
    st.dataframe(rfqs, use_container_width=True, hide_index=True)
    st.download_button("Download RFQs", to_excel_bytes(rfqs), "customer_rfqs.xlsx")

st.markdown(
    """
    <div class="footer-box">
        Ready Stock Parts | Industrial OEM spare parts inquiry portal for Saudi Arabia, GCC & Middle East.<br>
        Public pages do not display internal cost, total inventory value, or customer RFQ records.
    </div>
    """,
    unsafe_allow_html=True,
)
