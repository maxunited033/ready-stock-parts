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


def patch_streamlit_frontend() -> None:
    package_root = Path(streamlit.__file__).resolve().parent
    candidates = []

    # Known Streamlit locations across versions.
    known = [
        package_root / "static" / "index.html",
        package_root / "web" / "server" / "static" / "index.html",
    ]
    for path in known:
        if path.exists():
            candidates.append(path)

    # Fallback: search the installed package for any Streamlit HTML shell.
    for path in package_root.rglob("index.html"):
        if path not in candidates:
            try:
                sample = path.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                continue
            if "<title>Streamlit</title>" in sample or "window.prerenderReady" in sample:
                candidates.append(path)

    if not candidates:
        raise FileNotFoundError(
            f"No Streamlit frontend index.html found under {package_root}"
        )

    patched = []
    for index_file in candidates:
        html = index_file.read_text(encoding="utf-8", errors="ignore")

        updated = re.sub(
            r"<title>.*?</title>",
            f"<title>{TITLE}</title>",
            html,
            count=1,
            flags=re.IGNORECASE | re.DOTALL,
        )

        # Remove existing tags to avoid duplicates.
        updated = re.sub(
            r'<meta[^>]+name=["\']description["\'][^>]*>\s*',
            "",
            updated,
            flags=re.IGNORECASE,
        )
        updated = re.sub(
            r'<link[^>]+rel=["\']canonical["\'][^>]*>\s*',
            "",
            updated,
            flags=re.IGNORECASE,
        )
        updated = re.sub(
            r'<meta[^>]+property=["\']og:(?:site_name|title|description|type|url)["\'][^>]*>\s*',
            "",
            updated,
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

        updated = updated.replace("</head>", f"{seo_tags}</head>", 1)
        index_file.write_text(updated, encoding="utf-8")

        verification = index_file.read_text(encoding="utf-8", errors="ignore")
        if f"<title>{TITLE}</title>" not in verification:
            raise RuntimeError(f"Title patch verification failed for {index_file}")

        patched.append(str(index_file))

    print("Patched Streamlit frontend files:", flush=True)
    for item in patched:
        print(f" - {item}", flush=True)
    print(f"Verified title: {TITLE}", flush=True)


def start_streamlit() -> None:
    # Create the folder before Streamlit starts so static serving does not warn.
    Path("static").mkdir(parents=True, exist_ok=True)

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

    print("Starting Streamlit with:", " ".join(command), flush=True)
    os.execv(sys.executable, command)


if __name__ == "__main__":
    patch_streamlit_frontend()
    start_streamlit()
