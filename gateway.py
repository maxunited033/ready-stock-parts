import asyncio
import html
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from urllib.parse import urlencode, quote_plus

import httpx
import pandas as pd
import uvicorn
import websockets
from fastapi import FastAPI, Request, Response, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse, PlainTextResponse

SITE_URL = "https://www.readystockparts.com"
STREAMLIT_HOST = "127.0.0.1"
STREAMLIT_PORT = 8501
STREAMLIT_BASE = "/app"
INVENTORY_FILE = Path("data/inventory.xlsx")

app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
streamlit_process: subprocess.Popen | None = None


def load_inventory_index() -> dict[str, dict]:
    if not INVENTORY_FILE.exists():
        return {}
    try:
        df = pd.read_excel(INVENTORY_FILE)
    except Exception:
        return {}

    aliases = {
        "part": ["Part Number", "Name", "Internal Reference", "Item", "Item Code", "Part No", "Part No.", "P/N"],
        "description": ["Description", "Desc", "Product Description", "Item Description", "DESCRIPTION"],
        "manufacturer": ["Manufacturer", "OEM", "Brand", "Make", "Vendor"],
    }
    normalized = {str(c).strip().lower(): c for c in df.columns}

    def find_col(options):
        for item in options:
            if item.lower() in normalized:
                return normalized[item.lower()]
        return None

    part_col = find_col(aliases["part"])
    desc_col = find_col(aliases["description"])
    man_col = find_col(aliases["manufacturer"])
    if not part_col:
        return {}

    result = {}
    for _, row in df.iterrows():
        part = str(row.get(part_col, "")).strip()
        if not part or part.lower() == "nan":
            continue
        result[part.casefold()] = {
            "part": part,
            "description": str(row.get(desc_col, "Industrial spare part")).strip() if desc_col else "Industrial spare part",
            "manufacturer": str(row.get(man_col, "OEM")).strip() if man_col else "OEM",
        }
    return result


INVENTORY_INDEX = load_inventory_index()


def page_metadata(page: str, part: str = ""):
    page = (page or "home").strip().lower()
    part = (part or "").strip()

    if page == "part" and part:
        item = INVENTORY_INDEX.get(part.casefold(), {})
        clean_part = item.get("part", part)
        manufacturer = item.get("manufacturer", "Industrial")
        description = item.get("description", "Industrial spare part")
        title = f"{clean_part} | {manufacturer} Spare Parts | Ready Stock Parts"
        meta = (
            f"Request availability and pricing for part number {clean_part}. "
            f"{manufacturer} industrial spare parts support for Saudi Arabia and GCC customers."
        )
        heading = clean_part
        summary = f"{description} — {manufacturer}. Submit an RFQ for availability, price, and delivery."
        canonical = f"{SITE_URL}/?page=part&part={quote_plus(clean_part)}"
        return title, meta, heading, summary, canonical

    pages = {
        "home": (
            "Ready Stock Parts | Industrial OEM Spare Parts Supplier",
            "Search ready-stock OEM and aftermarket industrial spare parts and submit RFQs across Saudi Arabia and GCC.",
            "Ready Stock Parts",
            "Industrial OEM and aftermarket spare parts availability and RFQ platform.",
            f"{SITE_URL}/",
        ),
        "search": (
            "Search Industrial Spare Parts | Ready Stock Parts",
            "Search industrial spare parts by part number, description, manufacturer, OEM, or category.",
            "Search Available Spare Parts",
            "Search current industrial spare parts availability by part number, OEM, manufacturer, or description.",
            f"{SITE_URL}/?page=search",
        ),
        "rfq": (
            "Request a Spare Parts Quotation | Ready Stock Parts",
            "Submit single-line or multi-line RFQs for industrial OEM and aftermarket spare parts.",
            "Request for Quotation",
            "Submit one or multiple industrial spare parts in a single RFQ.",
            f"{SITE_URL}/?page=rfq",
        ),
        "brands": (
            "OEM Brands and Manufacturers | Ready Stock Parts",
            "Browse industrial OEM and manufacturer coverage available through Ready Stock Parts.",
            "OEM and Manufacturer Coverage",
            "Browse industrial manufacturers represented in the current ready-stock inventory.",
            f"{SITE_URL}/?page=brands",
        ),
        "about": (
            "About Ready Stock Parts | Saudi Arabia and GCC",
            "Learn about Ready Stock Parts and our industrial spare parts support for Saudi Arabia and GCC customers.",
            "About Ready Stock Parts",
            "B2B industrial spare parts availability and quotation support across Saudi Arabia and GCC.",
            f"{SITE_URL}/?page=about",
        ),
        "contact": (
            "Contact Ready Stock Parts | Sales and RFQ Support",
            "Contact Ready Stock Parts for urgent RFQs, spare parts identification, and availability checks.",
            "Contact Ready Stock Parts",
            "Contact our sales team for urgent RFQs and industrial spare parts identification.",
            f"{SITE_URL}/?page=contact",
        ),
        "privacy": (
            "Privacy Policy | Ready Stock Parts",
            "Read the Ready Stock Parts privacy policy for website visitors and RFQ customers.",
            "Privacy Policy",
            "How Ready Stock Parts handles customer and RFQ information.",
            f"{SITE_URL}/?page=privacy",
        ),
        "terms": (
            "Terms of Use | Ready Stock Parts",
            "Read the commercial website terms and conditions for Ready Stock Parts.",
            "Terms of Use",
            "Commercial terms governing use of the Ready Stock Parts website.",
            f"{SITE_URL}/?page=terms",
        ),
        "rfq_policy": (
            "RFQ Policy | Ready Stock Parts",
            "Learn how Ready Stock Parts processes industrial spare parts quotation requests.",
            "RFQ Policy",
            "How quotation requests, files, and commercial responses are processed.",
            f"{SITE_URL}/?page=rfq_policy",
        ),
    }
    return pages.get(page, pages["home"])


def shell_html(request: Request) -> str:
    page = request.query_params.get("page", "home")
    part = request.query_params.get("part", "")
    title, description, heading, summary, canonical = page_metadata(page, part)

    query = dict(request.query_params)
    iframe_query = urlencode(query)
    iframe_url = f"{STREAMLIT_BASE}/"
    if iframe_query:
        iframe_url += f"?{iframe_query}"

    schema = {
        "@context": "https://schema.org",
        "@type": "Organization",
        "name": "Ready Stock Parts",
        "url": SITE_URL,
        "email": "sales@readystockparts.com",
        "telephone": "+966561261005",
        "address": {
            "@type": "PostalAddress",
            "addressLocality": "Dammam",
            "addressCountry": "SA",
        },
    }
    if page == "part" and part:
        item = INVENTORY_INDEX.get(part.casefold(), {})
        schema = {
            "@context": "https://schema.org",
            "@type": "Product",
            "name": item.get("part", part),
            "sku": item.get("part", part),
            "description": item.get("description", description),
            "brand": {"@type": "Brand", "name": item.get("manufacturer", "OEM")},
            "url": canonical,
            "offers": {
                "@type": "Offer",
                "availability": "https://schema.org/InStock",
                "url": canonical,
                "priceCurrency": "SAR",
                "description": "Price available by quotation",
            },
        }

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">
  <title>{html.escape(title)}</title>
  <meta name=\"description\" content=\"{html.escape(description, quote=True)}\">
  <link rel=\"canonical\" href=\"{html.escape(canonical, quote=True)}\">
  <meta name=\"robots\" content=\"index,follow,max-image-preview:large\">
  <meta property=\"og:site_name\" content=\"Ready Stock Parts\">
  <meta property=\"og:title\" content=\"{html.escape(title, quote=True)}\">
  <meta property=\"og:description\" content=\"{html.escape(description, quote=True)}\">
  <meta property=\"og:type\" content=\"website\">
  <meta property=\"og:url\" content=\"{html.escape(canonical, quote=True)}\">
  <meta name=\"twitter:card\" content=\"summary\">
  <meta name=\"twitter:title\" content=\"{html.escape(title, quote=True)}\">
  <meta name=\"twitter:description\" content=\"{html.escape(description, quote=True)}\">
  <script type=\"application/ld+json\">{json.dumps(schema, ensure_ascii=False)}</script>
  <style>
    html,body{{margin:0;width:100%;height:100%;overflow:hidden;background:#fff}}
    .seo-content{{position:absolute;left:-9999px;width:1px;height:1px;overflow:hidden}}
    iframe{{display:block;border:0;width:100%;height:100vh;background:#fff}}
  </style>
</head>
<body>
  <main class=\"seo-content\">
    <h1>{html.escape(heading)}</h1>
    <p>{html.escape(summary)}</p>
    <p>Ready Stock Parts serves industrial customers across Saudi Arabia and GCC countries.</p>
  </main>
  <iframe src=\"{html.escape(iframe_url, quote=True)}\" title=\"Ready Stock Parts\" loading=\"eager\"></iframe>
</body>
</html>"""


async def wait_for_streamlit(timeout: float = 45.0):
    deadline = time.time() + timeout
    url = f"http://{STREAMLIT_HOST}:{STREAMLIT_PORT}{STREAMLIT_BASE}/_stcore/health"
    async with httpx.AsyncClient() as client:
        while time.time() < deadline:
            try:
                response = await client.get(url, timeout=2)
                if response.status_code < 500:
                    return
            except Exception:
                pass
            await asyncio.sleep(0.5)
    raise RuntimeError("Streamlit did not become ready in time.")


@app.on_event("startup")
async def startup_event():
    global streamlit_process
    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "app.py",
        "--server.address=127.0.0.1",
        f"--server.port={STREAMLIT_PORT}",
        "--server.headless=true",
        f"--server.baseUrlPath={STREAMLIT_BASE.strip('/')}",
        "--server.enableCORS=false",
        "--server.enableXsrfProtection=false",
        "--server.fileWatcherType=none",
        "--browser.gatherUsageStats=false",
    ]
    streamlit_process = subprocess.Popen(command)
    await wait_for_streamlit()


@app.on_event("shutdown")
async def shutdown_event():
    global streamlit_process
    if streamlit_process and streamlit_process.poll() is None:
        streamlit_process.terminate()
        try:
            streamlit_process.wait(timeout=8)
        except subprocess.TimeoutExpired:
            streamlit_process.kill()


@app.get("/robots.txt")
async def robots():
    return PlainTextResponse(f"User-agent: *\nAllow: /\nSitemap: {SITE_URL}/sitemap.xml\n")


@app.get("/sitemap.xml")
async def sitemap():
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
    urls += [f"{SITE_URL}/?page=part&part={quote_plus(item['part'])}" for item in INVENTORY_INDEX.values()]
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    for url in urls:
        escaped = url.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        lines.append(f"<url><loc>{escaped}</loc></url>")
    lines.append("</urlset>")
    return Response("\n".join(lines), media_type="application/xml")


@app.get("/health")
async def health():
    return PlainTextResponse("ok")


@app.websocket("/app/{path:path}")
async def websocket_proxy(websocket: WebSocket, path: str):
    """
    Proxy Streamlit's websocket connection. The loading skeleton remains on
    screen whenever this connection fails, so this route must preserve binary
    and text frames in both directions.
    """
    requested_subprotocols = websocket.scope.get("subprotocols", [])
    selected_subprotocol = requested_subprotocols[0] if requested_subprotocols else None
    await websocket.accept(subprotocol=selected_subprotocol)

    target = f"ws://{STREAMLIT_HOST}:{STREAMLIT_PORT}{STREAMLIT_BASE}/{path}"
    if websocket.url.query:
        target += f"?{websocket.url.query}"

    connect_kwargs = {
        "open_timeout": 20,
        "ping_interval": 20,
        "ping_timeout": 20,
        "max_size": None,
    }
    if requested_subprotocols:
        connect_kwargs["subprotocols"] = requested_subprotocols

    try:
        async with websockets.connect(target, **connect_kwargs) as upstream:
            async def client_to_upstream():
                while True:
                    message = await websocket.receive()
                    message_type = message.get("type")

                    if message_type == "websocket.disconnect":
                        await upstream.close()
                        return

                    if message.get("bytes") is not None:
                        await upstream.send(message["bytes"])
                    elif message.get("text") is not None:
                        await upstream.send(message["text"])

            async def upstream_to_client():
                try:
                    async for message in upstream:
                        if isinstance(message, bytes):
                            await websocket.send_bytes(message)
                        else:
                            await websocket.send_text(message)
                finally:
                    try:
                        await websocket.close()
                    except Exception:
                        pass

            tasks = [
                asyncio.create_task(client_to_upstream()),
                asyncio.create_task(upstream_to_client()),
            ]
            done, pending = await asyncio.wait(
                tasks,
                return_when=asyncio.FIRST_COMPLETED,
            )
            for task in pending:
                task.cancel()
            for task in done:
                exc = task.exception()
                if exc:
                    raise exc

    except (WebSocketDisconnect, websockets.ConnectionClosed):
        pass
    except Exception as exc:
        print(f"WebSocket proxy error for {path}: {exc}", flush=True)
        try:
            await websocket.close(code=1011)
        except Exception:
            pass


@app.api_route("/app/{path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def http_proxy(request: Request, path: str):
    target = f"http://{STREAMLIT_HOST}:{STREAMLIT_PORT}{STREAMLIT_BASE}/{path}"
    if request.url.query:
        target += f"?{request.url.query}"
    excluded = {"host", "content-length", "connection", "upgrade"}
    headers = {k: v for k, v in request.headers.items() if k.lower() not in excluded}
    body = await request.body()
    async with httpx.AsyncClient(follow_redirects=False, timeout=60) as client:
        upstream = await client.request(request.method, target, headers=headers, content=body)
    response_headers = {k: v for k, v in upstream.headers.items() if k.lower() not in {"content-encoding", "transfer-encoding", "connection", "content-length"}}
    return Response(content=upstream.content, status_code=upstream.status_code, headers=response_headers, media_type=upstream.headers.get("content-type"))


@app.get("/{path:path}", response_class=HTMLResponse)
async def html_shell(request: Request, path: str):
    return HTMLResponse(shell_html(request))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    uvicorn.run("gateway:app", host="0.0.0.0", port=port, log_level="info")
