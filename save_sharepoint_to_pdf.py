"""
SharePoint Site Pages → PDF Exporter  (Universal Edition)
==========================================================
Works with ANY SharePoint site. Discovers all pages automatically.

SETUP (run once):
    pip install playwright
    playwright install chromium

USAGE:
    # You'll be prompted for the site URL:
    python save_sharepoint_to_pdf.py

    # Or pass the site URL directly:
    python save_sharepoint_to_pdf.py https://tenant.sharepoint.com/sites/MySite

PDFs are saved to a folder named after the site, next to this script.
Re-running skips pages already exported (safe to resume if interrupted).
"""

import asyncio
import re
import sys
import json
from pathlib import Path
from urllib.parse import urlparse, quote


# ── Helpers ───────────────────────────────────────────────────────────────────

def safe_filename(name: str, index: int) -> str:
    stem = name.replace(".aspx", "")
    stem = re.sub(r'[<>:"/\\|?*]', '_', stem).strip(". ")
    return f"{index + 1:03d}_{stem}.pdf"


def parse_site_info(url: str):
    """
    Extract the site base URL and REST API root from a SharePoint URL.
    Handles formats like:
      https://tenant.sharepoint.com/sites/SiteName
      https://tenant.sharepoint.com/sites/SiteName/SubSite
      https://tenant.sharepoint.com/sites/SiteName/SubSite/SitePages/Forms/ByAuthor.aspx
    """
    parsed = urlparse(url)
    # Walk path segments until we find the site root (stops before /SitePages/)
    parts = [p for p in parsed.path.split('/') if p]
    site_parts = []
    for part in parts:
        if part.lower() in ('sitepages', 'pages', 'lists', 'shared documents', '_layouts'):
            break
        site_parts.append(part)
    site_path = '/' + '/'.join(site_parts) if site_parts else ''
    base = f"{parsed.scheme}://{parsed.netloc}"
    site_url = base + site_path
    api_root = site_path + '/_api'
    return base, site_url, api_root


async def discover_pages(page, api_root: str):
    """Use the SharePoint REST API (via authenticated browser) to list all site pages."""
    endpoint = f"{api_root}/web/lists/getbytitle('Site%20Pages')/items?$select=Title,FileLeafRef,FileRef&$top=500&$orderby=Modified asc"
    result = await page.evaluate(f"""
        async () => {{
            const r = await fetch({json.dumps(endpoint)}, {{
                headers: {{ 'Accept': 'application/json;odata=verbose' }}
            }});
            const d = await r.json();
            return d.d ? d.d.results.map(p => ({{
                title: p.Title || p.FileLeafRef,
                file: p.FileLeafRef,
                ref: p.FileRef
            }})) : [];
        }}
    """)
    return result


# ── Main ──────────────────────────────────────────────────────────────────────

async def run(site_url: str):
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("\n❌  Playwright is not installed.")
        print("    Run:  pip install playwright && playwright install chromium\n")
        sys.exit(1)

    base_url, site_url, api_root = parse_site_info(site_url)

    # Output folder named after the site
    site_name = site_url.rstrip('/').split('/')[-1] or 'SharePoint'
    output_dir = Path(__file__).parent / f"{site_name}_PDFs"
    profile_dir = Path(__file__).parent / "chromium_profile"
    output_dir.mkdir(exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  SharePoint → PDF Exporter")
    print(f"{'='*60}")
    print(f"  Site     : {site_url}")
    print(f"  Output   : {output_dir}")
    print(f"{'='*60}\n")

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            headless=False,
            no_viewport=True,
            args=["--start-maximized", "--window-size=1600,1000"],
        )

        page = context.pages[0] if context.pages else await context.new_page()

        # ── Login ─────────────────────────────────────────────────────────────
        print("🌐  Opening SharePoint...")
        login_url = site_url + "/SitePages/Forms/ByAuthor.aspx"
        await page.goto(login_url, wait_until="commit", timeout=60_000)
        await page.wait_for_load_state("domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(2000)

        title = await page.title()
        needs_login = any(w in title.lower() for w in ("sign", "login", "microsoft", "authenticat"))
        if needs_login:
            print("\n🔐  Please log in to SharePoint in the browser window.")
            print("    Once you can see the page list, press Enter here...")
            input()
        else:
            print("✅  Already logged in.\n")

        # ── Discover pages ────────────────────────────────────────────────────
        print("🔍  Discovering site pages via REST API...")
        pages = await discover_pages(page, api_root)

        if not pages:
            print("⚠️  No pages found. The site URL may be wrong, or you may need to log in first.")
            await context.close()
            return

        print(f"    Found {len(pages)} pages.\n")

        already_done = {f.name for f in output_dir.glob("*.pdf")}
        remaining = [(i, p) for i, p in enumerate(pages)
                     if safe_filename(p['file'], i) not in already_done]

        if not remaining:
            print("✅  All pages already exported. Nothing to do.\n")
            await context.close()
            return

        print(f"  To export : {len(remaining)}  (skipping {len(pages) - len(remaining)} already done)\n")

        # ── Export loop ───────────────────────────────────────────────────────
        errors = []
        base_site = base_url

        for idx, (page_index, sp_page) in enumerate(remaining):
            filename = sp_page['file']
            file_ref  = sp_page['ref']
            pdf_name  = safe_filename(filename, page_index)
            pdf_path  = output_dir / pdf_name
            url = base_site + file_ref.replace(" ", "%20")

            print(f"[{idx + 1:>3}/{len(remaining)}] {filename[:70]}")

            try:
                await page.goto(url, wait_until="commit", timeout=60_000)
                await page.wait_for_load_state("domcontentloaded", timeout=30_000)
                await page.wait_for_timeout(4000)

                try:
                    await page.wait_for_selector(
                        "[data-automation-id='pageContent'], .ms-rtestate-field, article, #contentBox",
                        timeout=10_000
                    )
                except Exception:
                    pass

                await page.pdf(
                    path=str(pdf_path),
                    format="A4",
                    print_background=True,
                    margin={"top": "15mm", "bottom": "15mm",
                            "left": "15mm", "right": "15mm"},
                )
                print(f"         ✅  Saved: {pdf_name}")

            except Exception as e:
                msg = str(e)[:120]
                print(f"         ❌  Error: {msg}")
                errors.append((filename, msg))

        await context.close()

    # ── Summary ───────────────────────────────────────────────────────────────
    saved = len(remaining) - len(errors)
    print(f"\n{'='*60}")
    print(f"  Done!  {saved}/{len(remaining)} PDFs saved to:")
    print(f"  {output_dir}")
    if errors:
        print(f"\n  ⚠️  {len(errors)} failed:")
        for name, err in errors:
            print(f"     • {name}: {err}")
    print(f"{'='*60}\n")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        url = sys.argv[1].strip()
    else:
        print("\nSharePoint Site → PDF Exporter")
        print("Paste the URL of any SharePoint site (or a page within it):")
        url = input("  URL: ").strip()

    if not url.startswith("http"):
        print("❌  Invalid URL. Must start with https://")
        sys.exit(1)

    asyncio.run(run(url))
