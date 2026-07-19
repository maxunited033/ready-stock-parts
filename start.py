import os
import re
import sys
from pathlib import Path

import streamlit


TITLE = "Ready Stock Parts | Industrial OEM Spare Parts Supplier"
DESCRIPTION = (
    "Ready Stock Parts is a B2B industrial spare parts platform serving Saudi Arabia "
    "and GCC customers with OEM and aftermarket spare parts availability checks and RFQ submissions."
)
SITE_URL = "https://www.readystockparts.com/"


def patch_streamlit_index() -> None:
    """
    Patch Streamlit's static HTML shell BEFORE the Streamlit server starts.
    This ensures View Source and search-engine crawlers see Ready Stock Parts
    instead of the default <title>Streamlit</title>.
    """
    static_index = Path(streamlit.__file__).resolve().parent / "static" / "index.html"

    if not static_index.exists():
        raise FileNotFoundError(f"Streamlit index.html not found: {static_index}")

    html = static_index.read_text(encoding="utf-8")

    html = re.sub(
        r"<title>.*?</title>",
        f"<title>{TITLE}</title>",
        html,
        count=1,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # Remove existing description/OG tags if present to avoid duplicates.
    html = re.sub(
        r'<meta[^>]+name=["\']description["\'][^>]*>\s*',
        "",
        html,
        flags=re.IGNORECASE,
    )
    html = re.sub(
        r'<meta[^>]+property=["\']og:(?:site_name|title|description|type|url)["\'][^>]*>\s*',
        "",
        html,
        flags=re.IGNORECASE,
    )

    seo_tags = (
        f'<meta name="description" content="{DESCRIPTION}" />\n'
        f'<link rel="canonical" href="{SITE_URL}" />\n'
        f'<meta property="og:site_name" content="Ready Stock Parts" />\n'
        f'<meta property="og:title" content="{TITLE}" />\n'
        f'<meta property="og:description" content="{DESCRIPTION}" />\n'
        f'<meta property="og:type" content="website" />\n'
        f'<meta property="og:url" content="{SITE_URL}" />\n'
        f'<meta name="twitter:card" content="summary" />\n'
        f'<meta name="twitter:title" content="{TITLE}" />\n'
        f'<meta name="twitter:description" content="{DESCRIPTION}" />\n'
    )

    html = html.replace("</head>", f"{seo_tags}</head>", 1)
    static_index.write_text(html, encoding="utf-8")

    print(f"Patched Streamlit metadata in: {static_index}")
    print(f"Title: {TITLE}")


def start_streamlit() -> None:
    app_file = os.environ.get("STREAMLIT_APP_FILE", "app.py")
    port = os.environ.get("PORT", "10000")

    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        app_file,
        "--server.address=0.0.0.0",
        f"--server.port={port}",
        "--server.headless=true",
    ]

    os.execv(sys.executable, command)


if __name__ == "__main__":
    patch_streamlit_index()
    start_streamlit()
