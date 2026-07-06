import os
from datetime import datetime
from io import BytesIO
from urllib.parse import quote_plus
import json
import urllib.request
import urllib.error

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Ready Stock Parts | Industrial OEM Spare Parts",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_DIR = "data"
DEFAULT_INVENTORY = os.path.join(DATA_DIR, "inventory.xlsx")
RFQ_FILE = os.path.join(DATA_DIR, "customer_rfqs.xlsx")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "Manutd@033")
CONTACT_EMAIL = "mossab.rozi@gmail.com"
CONTACT_MOBILE = "+966561261005"
WHATSAPP_NUMBER = "966561261005"

# Email notification settings via Resend API. Set these in Render > Environment.
RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
RESEND_FROM_EMAIL = os.environ.get("RESEND_FROM_EMAIL", "Ready Stock Parts <sales@readystockparts.com>")
RFQ_TO_EMAIL = os.environ.get("RFQ_TO_EMAIL", CONTACT_EMAIL)

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
        section[data-testid="stSidebar"] {display: none !important;}
        div[data-testid="collapsedControl"] {display: none !important;}
        .rsp-navbar {
            position: sticky; top: 0; z-index: 999;
            background: rgba(255,255,255,0.96);
            backdrop-filter: blur(10px);
            border: 1px solid #e5e7eb;
            border-radius: 18px;
            padding: 0.75rem 1rem;
            margin-bottom: 1rem;
            box-shadow: 0 8px 28px rgba(15,23,42,0.06);
            display: flex; align-items: center; justify-content: space-between; gap: 1rem; flex-wrap: wrap;
        }
        .rsp-brand {font-weight: 900; color: #0f172a; font-size: 1.1rem; letter-spacing: -0.02em;}
        .rsp-brand span {color: #2563eb;}
        .rsp-links {display:flex; gap:0.35rem; flex-wrap:wrap; align-items:center;}
        .rsp-links a {
            color:#334155 !important; text-decoration:none !important;
            padding:0.5rem 0.75rem; border-radius:999px; font-weight:650; font-size:0.92rem;
        }
        .rsp-links a:hover {background:#eff6ff; color:#1d4ed8 !important;}
        .rsp-links a.active {background:#1e3a8a; color:white !important;}
        .rsp-cta {background:#0f172a !important; color:white !important;}
        @media (max-width: 760px) {
            .block-container {padding-left: 1rem; padding-right: 1rem;}
            .hero {padding: 1.5rem 1.2rem; border-radius: 18px;}
            .hero h1 {font-size: 2rem;}
            .rsp-navbar {border-radius: 14px; padding: 0.65rem;}
            .rsp-links {width:100%; overflow-x:auto; flex-wrap:nowrap; padding-bottom:0.15rem;}
            .rsp-links a {white-space:nowrap; font-size:0.86rem;}
        }
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
            if len(df):
                lengths = df[col].fillna("").astype(str).str.len()
                q90 = lengths.quantile(0.9)
                if pd.isna(q90):
                    q90 = 14
                width = min(max(len(str(col)) + 4, int(q90) + 2), 48)
            else:
                width = max(len(str(col)) + 4, 14)
            ws.set_column(i, i, width)
    output.seek(0)
    return output.getvalue()


def send_rfq_email(row):
    """Send RFQ notification email through Resend API."""
    if not RESEND_API_KEY:
        return False, "Email notification is not configured. Add RESEND_API_KEY in Render Environment."
    if not RFQ_TO_EMAIL:
        return False, "RFQ recipient email is not configured. Add RFQ_TO_EMAIL in Render Environment."

    subject = f"New RFQ Received - {row.get('Part Number', '')}"
    customer_email = str(row.get("Email", "")).strip()

    text_body = f"""
New RFQ received from Ready Stock Parts website.

Date: {row.get('Date', '')}
Company: {row.get('Company', '')}
Contact Person: {row.get('Contact', '')}
Email: {row.get('Email', '')}
Mobile / WhatsApp: {row.get('Mobile', '')}
Part Number: {row.get('Part Number', '')}
Required Qty: {row.get('Required Qty', '')}
Notes: {row.get('Notes', '')}
Status: {row.get('Status', '')}

Website: https://readystockparts.com
"""

    html_body = f"""
    <div style="font-family: Arial, sans-serif; line-height:1.55; color:#0f172a;">
      <h2 style="margin-bottom:8px;">New RFQ Received</h2>
      <p>A new quotation request was submitted through <b>Ready Stock Parts</b>.</p>
      <table style="border-collapse:collapse; width:100%; max-width:720px;">
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Date</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Date', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Company</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Company', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Contact</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Contact', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Email</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Email', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Mobile / WhatsApp</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Mobile', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Part Number</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Part Number', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Required Qty</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Required Qty', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Notes</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Notes', '')}</td></tr>
      </table>
      <p style="margin-top:14px;"><a href="https://readystockparts.com">Open Ready Stock Parts</a></p>
    </div>
    """

    payload = {
        "from": RESEND_FROM_EMAIL,
        "to": [RFQ_TO_EMAIL],
        "subject": subject,
        "html": html_body,
        "text": text_body,
    }
    if customer_email and "@" in customer_email:
        payload["reply_to"] = customer_email

    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        "https://api.resend.com/emails",
        data=data,
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            if 200 <= response.status < 300:
                return True, "Email notification sent successfully via sales@readystockparts.com."
            response_body = response.read().decode("utf-8", errors="ignore")
            return False, f"RFQ saved, but Resend returned status {response.status}: {response_body}"
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="ignore")
        return False, f"RFQ saved, but email notification failed: {exc.code} {error_body}"
    except Exception as exc:
        return False, f"RFQ saved, but email notification failed: {exc}"


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
    with st.container():
        pwd = st.text_input("Admin Password", type="password", placeholder="Enter admin password")
    if pwd != ADMIN_PASSWORD:
        st.warning("Please enter the admin password to access this section.")
        st.stop()


PAGES = {
    "home": "Home",
    "search": "Search Parts",
    "rfq": "Request Quotation",
    "brands": "OEM Brands",
    "about": "About",
    "admin": "Admin Dashboard",
    "upload": "Upload Inventory",
    "inbox": "RFQ Inbox",
}

PAGE_TO_SLUG = {v: k for k, v in PAGES.items()}


def get_current_page():
    try:
        slug = st.query_params.get("page", "home")
    except Exception:
        slug = "home"
    if isinstance(slug, list):
        slug = slug[0] if slug else "home"
    return PAGES.get(slug, "Home")


def top_navigation():
    current_page = get_current_page()
    current_slug = PAGE_TO_SLUG.get(current_page, "home")
    links = []
    main_links = [
        ("home", "Home"),
        ("search", "Search Parts"),
        ("rfq", "Request RFQ"),
        ("brands", "OEM Brands"),
        ("about", "About"),
        ("admin", "Admin"),
    ]
    for slug, label in main_links:
        active = " active" if slug == current_slug else ""
        extra = " rsp-cta" if slug == "rfq" else ""
        links.append('<a class="{}{}" href="?page={}">{}</a>'.format(active, extra, slug, label))
    nav_html = """
        <div class="rsp-navbar">
            <div class="rsp-brand">Ready Stock <span>Parts</span></div>
            <div class="rsp-links">{links}</div>
        </div>
    """.format(links="".join(links))
    st.markdown(nav_html, unsafe_allow_html=True)
    return current_page


inject_css()

if not os.path.exists(DEFAULT_INVENTORY):
    st.error("Inventory file not found. Please make sure data/inventory.xlsx exists in the project.")
    st.stop()

inventory = load_inventory(DEFAULT_INVENTORY)
public_inventory = inventory[inventory["Qty"] > 0].copy()

page = top_navigation()

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
                rfq_row = {
                    "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Company": company,
                    "Contact": contact,
                    "Email": email,
                    "Mobile": mobile,
                    "Part Number": part,
                    "Required Qty": qty,
                    "Notes": notes,
                    "Status": "New",
                }
                save_rfq(rfq_row)
                email_sent, email_message = send_rfq_email(rfq_row)
                st.success("Your RFQ has been submitted successfully. We will contact you shortly.")
                if email_sent:
                    st.info(email_message)
                else:
                    st.warning(email_message)

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
    st.markdown("[Upload Inventory](?page=upload) &nbsp; | &nbsp; [RFQ Inbox](?page=inbox)", unsafe_allow_html=True)
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
