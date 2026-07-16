import os
from datetime import datetime
from io import BytesIO
from urllib.parse import quote_plus
from xml.sax.saxutils import escape as xml_escape
import json
import base64
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
CONTACT_EMAIL = "sales@readystockparts.com"
CONTACT_MOBILE = "+966561261005"
WHATSAPP_NUMBER = "966561261005"
SITE_URL = "https://www.readystockparts.com"

# Email notification settings via Resend API. Set these in Render > Environment.
RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
RESEND_FROM_EMAIL = os.environ.get("RESEND_FROM_EMAIL", "sales@readystockparts.com")
RFQ_TO_EMAIL = os.environ.get("RFQ_TO_EMAIL", "mossab.rozi@gmail.com")

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
            background: #ffffff !important; box-shadow: 0 8px 28px rgba(15,23,42,0.06); height: 100%;
            color: #0f172a !important;
        }
        .card h3 {margin-top: 0; color: #0f172a !important;}
        .card p, .card b, .card span {color: #0f172a !important;}
        .card .muted {color: #475569 !important;}
        .card a {color: #0f766e !important; font-weight: 700;}
        .contact-card {
            background: #ffffff !important;
            border: 1px solid #dbeafe !important;
            box-shadow: 0 12px 32px rgba(15,23,42,0.10) !important;
        }
        .contact-card p {font-size: 1rem;}
        .contact-card .contact-line {margin: 0.72rem 0;}
        .contact-card .contact-label {color: #475569 !important; font-weight: 700;}
        .contact-card .contact-value {color: #0f172a !important; font-weight: 800;}
        .muted {color: #64748b;}
        .section-title {font-size: 1.4rem; font-weight: 800; color: #0f172a; margin: 1.2rem 0 0.8rem;}
        .brand-pill {
            display: inline-block; padding: 0.55rem 0.85rem; border: 1px solid #dbeafe;
            border-radius: 999px; margin: 0.22rem; background: #eff6ff; color: #1e3a8a; font-weight: 650;
        }
        .footer-box {
            margin-top: 2.5rem;
            padding: 2rem;
            border: 1px solid #dbe3ee;
            border-radius: 22px;
            background: #0f172a !important;
            color: #e2e8f0 !important;
            box-shadow: 0 14px 40px rgba(15,23,42,0.16);
        }
        .footer-grid {
            display: grid;
            grid-template-columns: 1.45fr 1fr 1fr 1.2fr;
            gap: 1.8rem;
            align-items: start;
        }
        .footer-title {
            color: #ffffff !important;
            font-size: 1.05rem;
            font-weight: 900;
            margin-bottom: 0.8rem;
        }
        .footer-brand {
            color: #ffffff !important;
            font-size: 1.35rem;
            font-weight: 900;
            margin-bottom: 0.45rem;
        }
        .footer-tagline {
            color: #cbd5e1 !important;
            line-height: 1.65;
            margin-bottom: 0.85rem;
        }
        .footer-box a {
            color: #bfdbfe !important;
            text-decoration: none !important;
            display: block;
            margin: 0.42rem 0;
            font-weight: 650;
        }
        .footer-box a:hover {color: #ffffff !important; text-decoration: underline !important;}
        .footer-contact-line {color: #e2e8f0 !important; margin: 0.48rem 0;}
        .footer-trust {
            margin-top: 1.25rem;
            padding-top: 1.05rem;
            border-top: 1px solid rgba(255,255,255,0.14);
            color: #cbd5e1 !important;
            text-align: center;
            font-weight: 700;
        }
        .footer-bottom {
            margin-top: 0.85rem;
            text-align: center;
            color: #94a3b8 !important;
            font-size: 0.9rem;
        }
        .policy-card {
            border: 1px solid #dbe3ee;
            border-radius: 20px;
            padding: 1.5rem;
            background: #ffffff !important;
            color: #0f172a !important;
            box-shadow: 0 10px 30px rgba(15,23,42,0.07);
            margin-bottom: 1rem;
        }
        .policy-card h2, .policy-card h3, .policy-card p, .policy-card li {
            color: #0f172a !important;
        }
        .policy-card .policy-muted {color: #475569 !important;}
        @media (max-width: 900px) {
            .footer-grid {grid-template-columns: 1fr 1fr;}
        }
        @media (max-width: 620px) {
            .footer-grid {grid-template-columns: 1fr;}
            .footer-box {padding: 1.35rem;}
        }
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
            padding:0.55rem 0.85rem; border-radius:999px; font-weight:700; font-size:0.94rem;
        }
        .rsp-links a:hover {background:#eff6ff; color:#1d4ed8 !important;}
        .rsp-links a.active {background:#1e3a8a; color:#ffffff !important;}
        .rsp-cta {
            background: linear-gradient(135deg, #2563eb 0%, #0ea5e9 100%) !important;
            color:#ffffff !important;
            border: 1px solid rgba(255,255,255,0.35) !important;
            box-shadow: 0 8px 22px rgba(37,99,235,0.35);
            padding: 0.68rem 1.05rem !important;
            font-weight: 900 !important;
        }
        .rsp-cta::before {content: "📄 ";}
        .rsp-cta:hover {
            background: linear-gradient(135deg, #1d4ed8 0%, #0284c7 100%) !important;
            color:#ffffff !important;
            transform: translateY(-1px);
        }
        div[data-testid="stHorizontalBlock"]:has(button[kind]) {
            gap: 0.35rem;
        }
        button[kind="primary"] {
            font-weight: 850 !important;
            border-radius: 999px !important;
        }
        button[kind="secondary"] {
            font-weight: 750 !important;
            border-radius: 999px !important;
        }
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


RFQ_COLUMNS = [
    "Date", "Company", "Contact", "Email", "Mobile", "Line Items", "Attachment Names",
    "Status", "Email Notification", "Email Details", "Resend ID"
]


def load_rfqs():
    if os.path.exists(RFQ_FILE):
        rfqs = pd.read_excel(RFQ_FILE)
        for col in RFQ_COLUMNS:
            if col not in rfqs.columns:
                rfqs[col] = ""
        return rfqs[RFQ_COLUMNS]
    return pd.DataFrame(columns=RFQ_COLUMNS)


def save_rfq(row):
    os.makedirs(DATA_DIR, exist_ok=True)
    rfqs = load_rfqs()
    rfqs = pd.concat([rfqs, pd.DataFrame([row])], ignore_index=True)
    for col in RFQ_COLUMNS:
        if col not in rfqs.columns:
            rfqs[col] = ""
    rfqs[RFQ_COLUMNS].to_excel(RFQ_FILE, index=False)


def parse_line_items(row):
    raw = row.get("Line Items", "")
    if isinstance(raw, str) and raw.strip():
        try:
            items = json.loads(raw)
            if isinstance(items, list):
                return items
        except Exception:
            return []
    return []


def line_items_to_text(items):
    if not items:
        return "No line items provided."
    lines = []
    for idx, item in enumerate(items, start=1):
        lines.append(
            f"{idx}. Part Number: {item.get('Part Number', '')} | "
            f"Qty: {item.get('Qty', '')} | "
            f"Manufacturer: {item.get('Manufacturer', '')} | "
            f"Notes: {item.get('Notes', '')}"
        )
    return "\n".join(lines)


def line_items_to_html_rows(items):
    if not items:
        return '<tr><td colspan="4" style="padding:8px; border:1px solid #e5e7eb;">No line items provided.</td></tr>'
    rows = []
    for idx, item in enumerate(items, start=1):
        rows.append(
            f"<tr>"
            f"<td style='padding:8px; border:1px solid #e5e7eb;'>{idx}</td>"
            f"<td style='padding:8px; border:1px solid #e5e7eb;'>{item.get('Part Number', '')}</td>"
            f"<td style='padding:8px; border:1px solid #e5e7eb;'>{item.get('Qty', '')}</td>"
            f"<td style='padding:8px; border:1px solid #e5e7eb;'>{item.get('Manufacturer', '')}</td>"
            f"<td style='padding:8px; border:1px solid #e5e7eb;'>{item.get('Notes', '')}</td>"
            f"</tr>"
        )
    return "".join(rows)


def build_resend_attachments(uploaded_files):
    attachments = []
    total_bytes = 0
    if not uploaded_files:
        return attachments, None
    for uploaded in uploaded_files:
        content = uploaded.getvalue()
        total_bytes += len(content)
        # Keep total RFQ attachments comfortably below common API limits.
        if total_bytes > 8 * 1024 * 1024:
            return [], "Attachments exceed 8 MB total. RFQ was saved, but attachments were not emailed."
        attachments.append({
            "filename": uploaded.name,
            "content": base64.b64encode(content).decode("utf-8"),
        })
    return attachments, None

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


def send_rfq_email(row, uploaded_files=None):
    """Send RFQ notification email through Resend API."""
    if not RESEND_API_KEY.strip():
        return False, "Email notification is not configured. Add RESEND_API_KEY in Render Environment.", ""
    if not RFQ_TO_EMAIL:
        return False, "RFQ recipient email is not configured. Add RFQ_TO_EMAIL in Render Environment.", ""

    items = parse_line_items(row)
    item_count = len(items)
    first_part = items[0].get("Part Number", "") if items else "No Part Number"
    subject = f"New RFQ Received - {row.get('Company', '')} - {item_count} Item(s) - {first_part}"
    customer_email = str(row.get("Email", "")).strip()

    text_body = f"""
New RFQ received from Ready Stock Parts website.

Date: {row.get('Date', '')}
Company: {row.get('Company', '')}
Contact Person: {row.get('Contact', '')}
Email: {row.get('Email', '')}
Mobile / WhatsApp: {row.get('Mobile', '')}
Attachment Names: {row.get('Attachment Names', '')}
Status: {row.get('Status', '')}

Line Items:
{line_items_to_text(items)}

Website: https://readystockparts.com
"""

    html_body = f"""
    <div style="font-family: Arial, sans-serif; line-height:1.55; color:#0f172a;">
      <h2 style="margin-bottom:8px;">New RFQ Received</h2>
      <p>A new quotation request was submitted through <b>Ready Stock Parts</b>.</p>
      <table style="border-collapse:collapse; width:100%; max-width:760px; margin-bottom:16px;">
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Date</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Date', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Company</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Company', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Contact</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Contact', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Email</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Email', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Mobile / WhatsApp</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Mobile', '')}</td></tr>
        <tr><td style="padding:8px; border:1px solid #e5e7eb;"><b>Attachments</b></td><td style="padding:8px; border:1px solid #e5e7eb;">{row.get('Attachment Names', '')}</td></tr>
      </table>
      <h3>Requested Line Items</h3>
      <table style="border-collapse:collapse; width:100%; max-width:900px;">
        <tr style="background:#f8fafc;">
          <th style="padding:8px; border:1px solid #e5e7eb; text-align:left;">#</th>
          <th style="padding:8px; border:1px solid #e5e7eb; text-align:left;">Part Number</th>
          <th style="padding:8px; border:1px solid #e5e7eb; text-align:left;">Qty</th>
          <th style="padding:8px; border:1px solid #e5e7eb; text-align:left;">Manufacturer</th>
          <th style="padding:8px; border:1px solid #e5e7eb; text-align:left;">Notes</th>
        </tr>
        {line_items_to_html_rows(items)}
      </table>
      <p style="margin-top:14px;"><a href="https://readystockparts.com">Open Ready Stock Parts</a></p>
    </div>
    """

    attachments, attachment_warning = build_resend_attachments(uploaded_files or [])

    payload = {
        "from": RESEND_FROM_EMAIL,
        "to": [RFQ_TO_EMAIL],
        "subject": subject,
        "html": html_body,
        "text": text_body,
    }
    if attachments:
        payload["attachments"] = attachments
    if customer_email and "@" in customer_email:
        payload["reply_to"] = customer_email

    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        "https://api.resend.com/emails",
        data=data,
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY.strip()}",
            "Content-Type": "application/json",
            "User-Agent": "ReadyStockParts/1.0",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=25) as response:
            response_body = response.read().decode("utf-8", errors="ignore")
            resend_id = ""
            try:
                resend_id = json.loads(response_body).get("id", "")
            except Exception:
                pass
            if 200 <= response.status < 300:
                msg = "Email notification sent successfully via sales@readystockparts.com."
                if attachment_warning:
                    msg += " " + attachment_warning
                return True, msg, resend_id
            return False, f"Resend returned status {response.status}: {response_body}", resend_id
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="ignore")
        return False, f"HTTP {exc.code}: {error_body}", ""
    except Exception as exc:
        return False, str(exc), ""

def safe_query_value(name, default=""):
    try:
        value = st.query_params.get(name, default)
    except Exception:
        return default
    if isinstance(value, list):
        return str(value[0]) if value else default
    return str(value)


def part_page_url(part_number):
    return f"{SITE_URL}/?page=part&part={quote_plus(str(part_number).strip())}"


def write_seo_static_files(df):
    """Generate a sitemap from current in-stock inventory."""
    static_dir = Path("static")
    static_dir.mkdir(parents=True, exist_ok=True)

    urls = [
        f"{SITE_URL}/",
        f"{SITE_URL}/?page=search",
        f"{SITE_URL}/?page=rfq",
        f"{SITE_URL}/?page=brands",
        f"{SITE_URL}/?page=about",
        f"{SITE_URL}/?page=contact",
        f"{SITE_URL}/?page=privacy",
        f"{SITE_URL}/?page=terms",
        f"{SITE_URL}/?page=rfq_policy",
    ]

    unique_parts = (
        df["Part Number"]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s.ne("")]
        .drop_duplicates()
        .tolist()
    )
    urls.extend(part_page_url(part) for part in unique_parts)

    sitemap_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
    ]
    for url in urls:
        sitemap_lines.append("  <url>")
        sitemap_lines.append(f"    <loc>{xml_escape(url)}</loc>")
        sitemap_lines.append("  </url>")
    sitemap_lines.append("</urlset>")
    (static_dir / "sitemap.xml").write_text("\n".join(sitemap_lines), encoding="utf-8")

    robots = (
        "User-agent: *\n"
        "Allow: /\n"
        f"Sitemap: {SITE_URL}/app/static/sitemap.xml\n"
    )
    (static_dir / "robots.txt").write_text(robots, encoding="utf-8")


def inject_part_seo(part_number, manufacturer, description):
    """Update browser title and description for a part detail view."""
    safe_part = str(part_number).replace('"', "").replace("<", "").replace(">", "")
    safe_manufacturer = str(manufacturer).replace('"', "").replace("<", "").replace(">", "")
    safe_description = str(description).replace('"', "").replace("<", "").replace(">", "")
    title = f"{safe_part} | {safe_manufacturer} Spare Parts | Ready Stock Parts"
    meta_description = (
        f"Request a quotation for part number {safe_part}. "
        f"{safe_manufacturer} industrial spare parts availability for Saudi Arabia and GCC."
    )
    canonical = part_page_url(safe_part)

    st.html(
        f"""
        <script>
        document.title = {json.dumps(title)};
        let meta = document.querySelector('meta[name="description"]');
        if (!meta) {{
            meta = document.createElement('meta');
            meta.setAttribute('name', 'description');
            document.head.appendChild(meta);
        }}
        meta.setAttribute('content', {json.dumps(meta_description)});

        let canonical = document.querySelector('link[rel="canonical"]');
        if (!canonical) {{
            canonical = document.createElement('link');
            canonical.setAttribute('rel', 'canonical');
            document.head.appendChild(canonical);
        }}
        canonical.setAttribute('href', {json.dumps(canonical)});
        </script>
        """
    )


def render_part_detail(inventory_df):
    requested_part = safe_query_value("part", "").strip()
    if not requested_part:
        st.warning("No part number was selected.")
        if st.button("Go to Search Parts", type="primary"):
            st.query_params["page"] = "search"
            st.rerun()
        return

    match = inventory_df[
        inventory_df["Part Number"].astype(str).str.strip().str.casefold()
        == requested_part.casefold()
    ]

    if match.empty:
        st.error("This part number was not found in the current public inventory.")
        st.markdown(f"**Requested Part Number:** `{requested_part}`")
        if st.button("Submit a Manual RFQ", type="primary"):
            st.query_params["page"] = "rfq"
            st.query_params["part"] = requested_part
            st.rerun()
        return

    row = match.iloc[0]
    part_number = str(row.get("Part Number", requested_part))
    manufacturer = str(row.get("Manufacturer", "TBA"))
    description = str(row.get("Description", "TBA"))
    category = str(row.get("Category", "TBA"))
    availability = str(row.get("Availability", "Available on Request"))

    inject_part_seo(part_number, manufacturer, description)

    st.markdown(f"# {part_number}")
    st.caption("Industrial spare part availability and quotation request")

    c1, c2 = st.columns([1.55, 1])
    with c1:
        st.markdown("### Part Information")
        st.markdown(f"**Part Number:** `{part_number}`")
        st.markdown(f"**Manufacturer / OEM:** {manufacturer}")
        st.markdown(f"**Description:** {description}")
        st.markdown(f"**Category:** {category}")
        st.markdown(f"**Availability:** {availability}")
        st.info(
            "Pricing, delivery, and final stock confirmation are provided through an official quotation."
        )

    with c2:
        st.markdown("### Request Pricing")
        if st.button("Request RFQ for This Part", type="primary", use_container_width=True):
            st.query_params["page"] = "rfq"
            st.query_params["part"] = part_number
            st.rerun()

        whatsapp_text = quote_plus(
            f"Hello Mossab, I would like to request availability and pricing for part number {part_number}."
        )
        whatsapp_url = f"https://wa.me/{WHATSAPP_NUMBER}?text={whatsapp_text}"
        st.link_button(
            "WhatsApp Inquiry",
            whatsapp_url,
            use_container_width=True,
        )

        st.markdown(f"[Email Sales](mailto:{CONTACT_EMAIL}?subject=RFQ%20for%20{quote_plus(part_number)})")

    st.markdown("---")
    st.markdown("### Related Available Parts")
    related = inventory_df[
        (inventory_df["Manufacturer"].astype(str) == manufacturer)
        & (inventory_df["Part Number"].astype(str) != part_number)
    ].head(8).copy()

    if related.empty:
        st.caption("No related parts are currently displayed.")
    else:
        related["View"] = related["Part Number"].apply(part_page_url)
        st.dataframe(
            related[["Part Number", "Description", "Manufacturer", "Availability", "View"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "View": st.column_config.LinkColumn(
                    "Part Page",
                    display_text="Open",
                )
            },
        )


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
        <div class="card contact-card">
            <h3>Contact Ready Stock Parts</h3>
            <p class="muted">For urgent requirements, RFQs, or technical spare parts identification.</p>
            <p class="contact-line"><span class="contact-label">Email:</span> <span class="contact-value">{CONTACT_EMAIL}</span></p>
            <p class="contact-line"><span class="contact-label">Mobile / WhatsApp:</span> <span class="contact-value">{CONTACT_MOBILE}</span></p>
            <p><a href="{whatsapp_url}" target="_blank">💬 Open WhatsApp Inquiry</a></p>
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
    "contact": "Contact Us",
    "part": "Part Details",
    "privacy": "Privacy Policy",
    "terms": "Terms of Use",
    "rfq_policy": "RFQ Policy",
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


def navigate_to(slug):
    st.query_params["page"] = slug
    st.rerun()


def top_navigation():
    current_page = get_current_page()
    current_slug = PAGE_TO_SLUG.get(current_page, "home")

    st.markdown('<div class="rsp-navbar-anchor"></div>', unsafe_allow_html=True)
    cols = st.columns([1.25, 1.15, 1.35, 1.15, 0.9, 0.95, 0.85])

    nav_items = [
        ("home", "Home"),
        ("search", "Search Parts"),
        ("rfq", "Request RFQ"),
        ("brands", "OEM Brands"),
        ("about", "About"),
        ("contact", "Contact"),
        ("admin", "Admin"),
    ]

    for col, (slug, label) in zip(cols, nav_items):
        with col:
            button_type = "primary" if slug == current_slug or slug == "rfq" else "secondary"
            if st.button(
                label,
                key=f"nav_{slug}",
                use_container_width=True,
                type=button_type,
            ):
                navigate_to(slug)

    return current_page


inject_css()

if not os.path.exists(DEFAULT_INVENTORY):
    st.error("Inventory file not found. Please make sure data/inventory.xlsx exists in the project.")
    st.stop()

inventory = load_inventory(DEFAULT_INVENTORY)
public_inventory = inventory[inventory["Qty"] > 0].copy()
write_seo_static_files(public_inventory)

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

    public_result = result[SAFE_COLUMNS].copy()
    public_result["View Part"] = public_result["Part Number"].apply(part_page_url)
    st.dataframe(
        public_result,
        use_container_width=True,
        hide_index=True,
        column_config={
            "View Part": st.column_config.LinkColumn(
                "Part Page",
                display_text="Open",
            )
        },
    )
    st.download_button(
        "Download Customer Stock List",
        to_excel_bytes(result[SAFE_COLUMNS]),
        "ready_stock_parts_customer_list.xlsx",
    )

    st.markdown("---")
    st.markdown("**Need pricing or delivery details?** Go to the Request Quotation page and submit the required part number and quantity.")

elif page == "Request Quotation":
    hero()
    st.subheader("Request for Quotation")
    st.caption("Submit multiple spare parts in one RFQ. You may also attach an Excel or PDF RFQ file.")

    if "rfq_item_count" not in st.session_state:
        st.session_state.rfq_item_count = 1

    with st.form("rfq_form"):
        st.markdown("### Customer Information")
        col1, col2 = st.columns(2)
        with col1:
            company = st.text_input("Company Name *")
            contact = st.text_input("Contact Person *")
            email = st.text_input("Email *")
        with col2:
            mobile = st.text_input("Mobile / WhatsApp")
            project_ref = st.text_input("Project / RFQ Reference")

        st.markdown("### RFQ Line Items")
        st.caption("Enter one or more part numbers and quantities. Empty rows will be ignored.")
        line_items = []
        for i in range(st.session_state.rfq_item_count):
            cols = st.columns([2.2, 0.8, 1.4, 2.2])
            with cols[0]:
                default_part = safe_query_value("part", "") if i == 0 else ""
                part_no = st.text_input(
                    f"Part Number {i + 1}" + (" *" if i == 0 else ""),
                    value=default_part,
                    key=f"rfq_part_{i}",
                )
            with cols[1]:
                qty = st.number_input(f"Qty {i + 1}", min_value=1, value=1, key=f"rfq_qty_{i}")
            with cols[2]:
                manufacturer = st.text_input(f"Manufacturer {i + 1}", key=f"rfq_manufacturer_{i}")
            with cols[3]:
                item_notes = st.text_input(f"Line Notes {i + 1}", key=f"rfq_item_notes_{i}")
            if str(part_no).strip():
                line_items.append({
                    "Part Number": str(part_no).strip(),
                    "Qty": int(qty),
                    "Manufacturer": str(manufacturer).strip(),
                    "Notes": str(item_notes).strip(),
                })

        add_part = st.form_submit_button("➕ Add Another Part", type="primary", use_container_width=True)
        st.caption("Add as many part numbers as needed. Empty rows will be ignored.")
        if add_part:
            st.session_state.rfq_item_count += 1
            st.rerun()

        reset_items = st.form_submit_button("Reset Line Items")
        if reset_items:
            st.session_state.rfq_item_count = 1
            st.rerun()

        overall_notes = st.text_area("Overall Notes / Equipment Details")
        uploaded_files = st.file_uploader(
            "Attach RFQ file (Excel or PDF)",
            type=["xlsx", "xls", "pdf"],
            accept_multiple_files=True,
            help="Optional. Attach customer RFQ files, drawings, or Excel lists. Keep total attachments below 8 MB."
        )
        submitted = st.form_submit_button("Submit RFQ", type="primary")

        if submitted:
            if not company or not contact or not email:
                st.error("Please complete Company Name, Contact Person, and Email.")
            elif not line_items:
                st.error("Please enter at least one Part Number.")
            else:
                if overall_notes:
                    line_items.append({
                        "Part Number": "Overall Notes",
                        "Qty": "",
                        "Manufacturer": "",
                        "Notes": overall_notes.strip(),
                    })
                attachment_names = ", ".join([f.name for f in uploaded_files]) if uploaded_files else ""
                rfq_row = {
                    "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Company": company,
                    "Contact": contact,
                    "Email": email,
                    "Mobile": mobile,
                    "Line Items": json.dumps(line_items, ensure_ascii=False),
                    "Attachment Names": attachment_names,
                    "Status": "New",
                    "Email Notification": "Pending",
                    "Email Details": "",
                    "Resend ID": "",
                }
                email_sent, email_message, resend_id = send_rfq_email(rfq_row, uploaded_files)
                rfq_row["Email Notification"] = "Sent" if email_sent else "Failed"
                rfq_row["Email Details"] = email_message
                rfq_row["Resend ID"] = resend_id
                save_rfq(rfq_row)
                st.success("Your RFQ has been submitted successfully. We will contact you shortly.")
                if not email_sent:
                    st.warning("Your RFQ was saved. Email notification did not complete, but our team can still see it in the RFQ Inbox.")

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


elif page == "Contact Us":
    hero()
    st.subheader("Contact Us")
    st.caption("Connect with Ready Stock Parts for urgent RFQs, availability checks, and industrial spare parts identification.")
    left, right = st.columns([1.25, 1])
    with left:
        st.markdown(
            f"""
            <div class="policy-card">
                <h2>Ready Stock Parts</h2>
                <p class="policy-muted"><b>OEM & Aftermarket Industrial Spare Parts Supplier</b></p>
                <p>📍 Dammam, Eastern Province, Saudi Arabia</p>
                <p>✉️ <a href="mailto:{CONTACT_EMAIL}">{CONTACT_EMAIL}</a></p>
                <p>📞 <a href="tel:{CONTACT_MOBILE}">{CONTACT_MOBILE}</a></p>
                <p>🌍 Serving industrial customers across Saudi Arabia and GCC countries.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        contact_box()

elif page == "Privacy Policy":
    hero()
    st.subheader("Privacy Policy")
    st.caption("Last updated: July 2026")
    st.write(
        "Ready Stock Parts respects the privacy of customers and website visitors. "
        "This policy explains how information submitted through the website is handled."
    )

    st.markdown("### Information We Collect")
    st.markdown(
        """
- Company name, contact name, email address, and telephone or WhatsApp number.
- Part numbers, quantities, equipment details, and quotation requirements.
- Files voluntarily uploaded with an RFQ, including Excel and PDF documents.
        """
    )

    st.markdown("### How We Use Information")
    st.markdown(
        """
- To review availability and prepare commercial quotations.
- To contact customers regarding their RFQs or technical inquiries.
- To improve website functionality, customer service, and inventory planning.
        """
    )

    st.markdown("### Information Sharing")
    st.write(
        "Customer information is not sold. Information may be shared only with relevant "
        "suppliers, logistics providers, or service partners when necessary to process a quotation or order."
    )

    st.markdown("### Data Security")
    st.write(
        "Reasonable technical and organizational measures are used to protect submitted information. "
        "However, no internet transmission method can be guaranteed to be completely secure."
    )

    st.markdown("### Contact")
    st.markdown(f"Questions regarding this policy may be sent to [{CONTACT_EMAIL}](mailto:{CONTACT_EMAIL}).")

elif page == "Terms of Use":
    hero()
    st.subheader("Terms of Use")
    st.caption("Last updated: July 2026")

    st.markdown("### Website Purpose")
    st.write(
        "This website is a B2B industrial spare parts inquiry and RFQ platform. "
        "Information shown on the website is provided for general commercial reference only."
    )

    st.markdown("### Quotations and Availability")
    st.markdown(
        """
- Website availability indicators are not a binding stock commitment.
- Prices, delivery dates, warranty terms, and commercial conditions are valid only when included in an official written quotation.
- All quotations remain subject to final stock confirmation and supplier approval.
        """
    )

    st.markdown("### Product and Trademark Information")
    st.write(
        "Product names, OEM names, trademarks, and logos remain the property of their respective owners. "
        "Their appearance on this website does not imply ownership by Ready Stock Parts."
    )

    st.markdown("### Customer Responsibility")
    st.write(
        "Customers are responsible for confirming part-number accuracy, equipment compatibility, "
        "required quantities, and technical specifications before placing an order."
    )

    st.markdown("### Limitation of Liability")
    st.write(
        "Ready Stock Parts is not liable for losses resulting from reliance on preliminary website information, "
        "incorrect customer data, or delays outside its reasonable control."
    )

    st.markdown("### Contact")
    st.markdown(f"For questions regarding these terms, contact [{CONTACT_EMAIL}](mailto:{CONTACT_EMAIL}).")

elif page == "RFQ Policy":
    hero()
    st.subheader("RFQ Policy")
    st.caption("How quotation requests are processed")

    st.markdown("### Submitting an RFQ")
    st.markdown(
        """
- Customers may submit one or multiple part numbers in a single RFQ.
- Excel and PDF attachments are accepted for larger requests.
- Clear part numbers, quantities, OEM information, and equipment details help us respond faster.
        """
    )

    st.markdown("### Response Time")
    st.write(
        "We aim to acknowledge RFQs promptly and normally provide an update within 1–2 business days. "
        "Complex technical or multi-line requests may require additional time."
    )

    st.markdown("### Commercial Status")
    st.markdown(
        """
- An RFQ submission is not a purchase order or contractual commitment.
- Availability, pricing, lead time, freight, and payment terms are confirmed only in the official quotation.
- We may request further technical or commercial information before issuing a quotation.
        """
    )

    st.markdown("### Confidentiality")
    st.write(
        "RFQ information and attachments are handled for quotation and order-support purposes. "
        "Sensitive information should be limited to what is necessary for the inquiry."
    )

    st.markdown("### Support")
    st.markdown(
        f"For urgent requirements, email [{CONTACT_EMAIL}](mailto:{CONTACT_EMAIL}) "
        "or contact us through WhatsApp."
    )

elif page == "Part Details":
    hero()
    render_part_detail(public_inventory)

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

def corporate_footer():
    whatsapp_text = quote_plus(
        "Hello Mossab, I would like to inquire about industrial spare parts availability."
    )
    whatsapp_url = f"https://wa.me/{WHATSAPP_NUMBER}?text={whatsapp_text}"

    # Keep the HTML on one continuous line so Streamlit does not interpret
    # indented sections as a code block.
    footer_html = (
        '<div class="footer-box">'
        '<div class="footer-grid">'
        '<div>'
        '<div class="footer-brand">Ready Stock Parts</div>'
        '<div class="footer-tagline">OEM &amp; Aftermarket Industrial Spare Parts Supplier '
        'supporting procurement, maintenance, and engineering teams across Saudi Arabia and the GCC.</div>'
        '<div class="footer-contact-line">📍 Dammam, Saudi Arabia</div>'
        f'<div class="footer-contact-line">✉️ {CONTACT_EMAIL}</div>'
        f'<div class="footer-contact-line">📞 {CONTACT_MOBILE}</div>'
        '</div>'
        '<div>'
        '<div class="footer-title">Quick Links</div>'
        '<a href="?page=home" target="_self">Home</a>'
        '<a href="?page=search" target="_self">Search Parts</a>'
        '<a href="?page=rfq" target="_self">Request RFQ</a>'
        '<a href="?page=brands" target="_self">OEM Brands</a>'
        '<a href="?page=about" target="_self">About Us</a>'
        '</div>'
        '<div>'
        '<div class="footer-title">Information</div>'
        '<a href="?page=privacy" target="_self">Privacy Policy</a>'
        '<a href="?page=terms" target="_self">Terms of Use</a>'
        '<a href="?page=rfq_policy" target="_self">RFQ Policy</a>'
        '<a href="?page=contact" target="_self">Contact Us</a>'
        '</div>'
        '<div>'
        '<div class="footer-title">Connect</div>'
        f'<a href="mailto:{CONTACT_EMAIL}">Email Sales</a>'
        f'<a href="{whatsapp_url}" target="_blank">WhatsApp Inquiry</a>'
        '<div class="footer-contact-line" style="margin-top:0.9rem;">'
        'Business inquiries and quotation requests are welcome.'
        '</div>'
        '</div>'
        '</div>'
        '<div class="footer-trust">Serving Industrial Customers Across Saudi Arabia &amp; GCC</div>'
        '<div class="footer-bottom">© 2026 Ready Stock Parts. All Rights Reserved.</div>'
        '</div>'
    )
    st.markdown(footer_html, unsafe_allow_html=True)


corporate_footer()
