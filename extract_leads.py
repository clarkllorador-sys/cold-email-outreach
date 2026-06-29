"""
extract_opentable.py
Scrapes restaurant listings from OpenTable UK for configured cities,
extracts contact emails from each restaurant's website, and writes
cleaned/deduplicated contacts to a Google Sheets tab.

Usage:
    python extract_opentable.py
    python extract_opentable.py --cities London Leicester Birmingham
    python extract_opentable.py --limit 50  # cap per city, for testing
    python extract_opentable.py --no-gsheets  # dry run, prints to console only
"""

import argparse
import logging
import re
import time
import json
import os
import subprocess
import tempfile
from datetime import datetime

from pathlib import Path

import urllib.parse

import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from playwright_stealth import Stealth
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────

# Add or remove cities here freely
DEFAULT_CITIES = [
    # Major cities
    "London",
    "Birmingham",
    "Manchester",
    "Leeds",
    "Liverpool",
    "Bristol",
    "Newcastle",
    "Sheffield",
    "Nottingham",
    "Leicester",
    # Secondary cities
    "Southampton",
    "Portsmouth",
    "Brighton",
    "Oxford",
    "Cambridge",
    "Coventry",
    "Reading",
    "Derby",
    "York",
    "Bath",
    "Exeter",
    "Plymouth",
    "Norwich",
    "Cheltenham",
    "Gloucester",
    "Chester",
    "Ipswich",
    "Northampton",
    "Milton Keynes",
    "Swindon",
    "Bournemouth",
    "Peterborough",
    "Middlesbrough",
    "Sunderland",
    "Hull",
    "Bradford",
    "Stoke-on-Trent",
    "Wolverhampton",
    "Blackpool",
    "Luton",
    "Guildford",
    "Winchester",
    "Salisbury",
    "Hereford",
    "Shrewsbury",
    "Carlisle",
    "Durham",
]

OPENTABLE_SEARCH_URL = "https://www.opentable.co.uk/s/"


GSHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CONTACTS_TAB = "contacts"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# Seconds to wait between restaurant website requests (be polite)
CRAWL_DELAY = 1.5

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Email extraction ──────────────────────────────────────────────────────────

EMAIL_REGEX = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}"
)

EXCLUDED_EMAIL_DOMAINS = {
    "sentry.io", "wixpress.com", "squarespace.com", "wordpress.com",
    "example.com", "domain.com", "email.com", "yourwebsite.com",
    "google.com", "facebook.com", "instagram.com", "twitter.com",
    "cookieinformation.com", "cookiebot.com", "onetrust.com",
    "2x-scaled.jpg",  # safety net, caught by _FILE_EXTENSIONS below too
}

# TLDs that are actually file extensions — not real email domains
_FILE_EXTENSION_TLDS = {
    "jpg", "jpeg", "png", "gif", "svg", "webp", "ico", "bmp", "avif",
    "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx",
    "css", "js", "json", "xml", "zip", "mp4", "mp3", "woff", "woff2",
}

# Local parts (before @) that are clearly placeholders or templates
_PLACEHOLDER_LOCAL_PARTS = {
    "your", "you", "tba", "test", "name", "email",
    "user", "username", "placeholder", "example",
}

CONTACT_PAGE_HINTS = ["/contact", "/contact-us", "/about", "/reach-us", "/enquiries"]


def is_valid_email(email: str) -> bool:
    local, _, domain = email.partition("@")
    local = local.lower()
    domain = domain.lower()
    # Excluded domains (cookie banners, tracking, etc.)
    if domain in EXCLUDED_EMAIL_DOMAINS:
        return False
    # Reject if TLD is a file extension (e.g. image@2x-scaled.jpg)
    tld = domain.rsplit(".", 1)[-1] if "." in domain else ""
    if tld in _FILE_EXTENSION_TLDS:
        return False
    # Reject placeholder/template local parts (Your@email.here, tba@tba.com, etc.)
    if local in _PLACEHOLDER_LOCAL_PARTS:
        return False
    if "noreply" in local or "no-reply" in local:
        return False
    return True


def extract_emails_from_html(html: str) -> list[str]:
    found = EMAIL_REGEX.findall(html)
    return [e for e in set(found) if is_valid_email(e)]


PHONE_REGEX = re.compile(
    r"(\+44\s?[\d\s\-\(\)]{9,}|0[\d\s\-\(\)]{10,})"
)

ADDRESS_HINTS = [
    r"\b[A-Z]{1,2}\d{1,2}[A-Z]?\s?\d[A-Z]{2}\b",  # UK postcode
]
ADDRESS_REGEX = re.compile("|".join(ADDRESS_HINTS))


def get_lead_tier(count) -> str:
    """Derives lead tier from location count.
    Independent = 1, Small chain = 2-5, Large chain = 6+
    """
    if count is None:
        return ""
    try:
        count = int(count)
    except (ValueError, TypeError):
        return ""
    if count == 1:
        return "Independent"
    if count <= 5:
        return "Small chain"
    return "Large chain"


def bdata_fetch_html(url: str, timeout: int = 90) -> str | None:
    """Fetches a URL via `bdata scrape -o tmpfile`, returns HTML string or None on failure."""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
            tmp_path = tmp.name
        r = subprocess.run(
            ["bdata", "scrape", url, "-f", "html", "-o", tmp_path],
            capture_output=True,
            text=True,
            timeout=timeout,
            env=os.environ.copy(),
        )
        if r.returncode == 0 and os.path.exists(tmp_path):
            html = Path(tmp_path).read_text(errors="replace")
            if len(html) > 500:
                return html
        log.debug("bdata_fetch_html failed for %s (rc=%d)", url, r.returncode)
    except subprocess.TimeoutExpired:
        log.debug("bdata_fetch_html timed out for %s", url)
    except FileNotFoundError:
        log.error("bdata CLI not found — install: sudo npm install -g @brightdata/cli && bdata login")
    except Exception as exc:
        log.debug("bdata_fetch_html error for %s: %s", url, exc)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
    return None


def extract_instagram_url(html: str) -> str | None:
    """
    Finds the restaurant's Instagram profile URL from page HTML.
    Looks for <a href="https://www.instagram.com/..."> links (typically in footer/nav).
    Filters out post, reel, story, and explore URLs — only returns profile URLs.
    Returns a normalised https://www.instagram.com/{handle}/ or None.
    """
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=re.compile(r"instagram\.com", re.I)):
        href = a.get("href", "").strip()
        try:
            parsed = urllib.parse.urlparse(href)
            path = parsed.path.strip("/")
            if not path:
                continue
            first_segment = path.split("/")[0].lower()
            # Skip non-profile paths
            if first_segment in {"p", "reel", "reels", "stories", "explore", "tv", "accounts", ""}:
                continue
            return f"https://www.instagram.com/{first_segment}/"
        except Exception:
            continue
    return None


def fetch_contact_details(page, website_url: str) -> dict:
    """
    Visits a restaurant website using Playwright (full JS rendering) and extracts
    emails and Instagram profile URL.
    Checks the given URL + common contact page paths, aggregates all unique emails found.
    Instagram URL is grabbed from the homepage only (it's always in the footer/nav).
    Restaurant sites are JS-rendered but have no bot protection — Playwright handles them fine.
    Returns a dict with keys: email (comma-separated if multiple), instagram_url.
    """
    result = {"email": None, "instagram_url": None}

    if not website_url:
        return result

    base = website_url.rstrip("/")
    urls_to_try = [website_url] + [base + hint for hint in CONTACT_PAGE_HINTS]

    all_emails: list[str] = []
    seen: set[str] = set()

    for idx, url in enumerate(urls_to_try):
        try:
            page.goto(url, wait_until="networkidle", timeout=30000)
            html = page.content()

            # Grab Instagram URL from homepage only (first URL)
            if idx == 0 and not result["instagram_url"]:
                ig_url = extract_instagram_url(html)
                if ig_url:
                    result["instagram_url"] = ig_url
                    log.info("  Instagram URL found: %s", ig_url)

            emails = extract_emails_from_html(html)
            new = [e for e in emails if e not in seen]
            if new:
                log.info("Emails found at %s: %s", url, new)
                all_emails.extend(new)
                seen.update(new)
            time.sleep(CRAWL_DELAY)
        except Exception as exc:
            log.debug("Playwright contact fetch failed for %s: %s", url, exc)

    if all_emails:
        result["email"] = ", ".join(all_emails)

    return result


# ── OpenTable scraping ────────────────────────────────────────────────────────

def scrape_opentable_city(page, city: str, limit: int | None = None) -> list[dict]:
    """
    Uses a Playwright page to scrape OpenTable UK search results for a city.
    Returns a list of raw restaurant dicts.
    """
    log.info("Scraping OpenTable for city: %s", city)
    results = []

    try:
        try:
            page.goto(
                f"{OPENTABLE_SEARCH_URL}?term={city}&covers=2",
                wait_until="domcontentloaded",
                timeout=60000,
            )
        except Exception as e:
            log.warning("Page load issue for %s (%s) — attempting to parse whatever loaded", city, type(e).__name__)

        # Give JS time to render restaurant cards
        page.wait_for_timeout(6000)

        # Dismiss cookie banner if present
        try:
            page.click("button:has-text('Accept'), button:has-text('Accept All'), #onetrust-accept-btn-handler", timeout=3000)
            page.wait_for_timeout(1000)
        except Exception:
            pass

        # Scroll to load more results
        for _ in range(5):
            page.keyboard.press("End")
            page.wait_for_timeout(1500)

        # Parse listings
        html = page.content()

        soup = BeautifulSoup(html, "html.parser")

        # OpenTable restaurant cards — selector may need updating if OT changes their markup
        cards = soup.select("[data-test='restaurant-card'], [class*='restaurantCard'], [class*='RestaurantCard']")

        if not cards:
            # Fallback: try to find any anchor with /restaurant/ in href
            cards = soup.find_all("a", href=re.compile(r"/restaurant/"))

        log.info("Found %d raw cards for %s", len(cards), city)

        seen_urls = set()
        for card in cards:
            if limit and len(results) >= limit:
                break

            # Name — h6 inside each card
            name_el = card.select_one("h6")
            name = name_el.get_text(strip=True) if name_el else None

            # OpenTable URL — link to /r/ path
            link = card.select_one("a[href*='/r/']")
            if not link:
                continue
            href = link["href"]
            opentable_url = (
                href if href.startswith("http")
                else f"https://www.opentable.co.uk{href}"
            )

            if opentable_url in seen_urls:
                continue
            seen_urls.add(opentable_url)

            # Rating — aria-label="X stars"
            rating_el = card.select_one("[aria-label*='stars']")
            rating = None
            if rating_el:
                m = re.search(r"([\d.]+)\s+stars", rating_el.get("aria-label", ""))
                rating = m.group(1) if m else None

            # Review count — link to #reviews, text like (12141)
            review_el = card.select_one("a[href*='#reviews']")
            review_count = review_el.get_text(strip=True).strip("()") if review_el else None

            # Price range — aria-label="Price: Moderate"
            price_el = card.select_one("[aria-label*='Price']")
            price_range = None
            if price_el:
                m = re.search(r"Price:\s*(.+)", price_el.get("aria-label", ""))
                price_range = m.group(1).strip() if m else None

            # Cuisine + location — data-test="cuisine-and-location", text: " • British • Soho"
            cuisine_loc_el = card.select_one("[data-test='cuisine-and-location']")
            cuisine, location = None, city
            if cuisine_loc_el:
                parts = [p.strip() for p in cuisine_loc_el.get_text().split("•") if p.strip()]
                cuisine = parts[0] if len(parts) > 0 else None
                location = parts[1] if len(parts) > 1 else city

            results.append({
                "restaurant_name": name,
                "opentable_url": opentable_url,
                "cuisine": cuisine,
                "location": location,
                "city": city,
                "rating": rating,
                "review_count": review_count,
                "price_range": price_range,
            })

    except PWTimeout:
        log.warning("Timeout scraping OpenTable for %s", city)
    except Exception as exc:
        log.error("Error scraping %s: %s", city, exc)

    log.info("Scraped %d listings for %s", len(results), city)
    return results


SKIP_DOMAINS = {
    "opentable", "tripadvisor", "google", "facebook", "instagram",
    "twitter", "yelp", "timeout", "squarespace", "wix", "wikipedia",
    "thefork", "bookatable", "designmynight", "eatout", "hardens",
}

# File extensions and subdomain patterns that are clearly not restaurant homepages
_BAD_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".csv"}
_BAD_SUBDOMAINS = {"downloads.", "cdn.", "assets.", "media.", "static.", "files."}


def is_restaurant_website(url: str) -> bool:
    """Returns False for URLs that are clearly not restaurant homepages (PDFs, download servers, etc.)."""
    if not url:
        return False
    try:
        parsed = urllib.parse.urlparse(url)
        if any(parsed.path.lower().endswith(ext) for ext in _BAD_EXTENSIONS):
            return False
        if any(parsed.netloc.lower().startswith(sub) for sub in _BAD_SUBDOMAINS):
            return False
    except Exception:
        return False
    return True


def _extract_from_ot_json(html: str, restaurant: dict) -> bool:
    """
    Parses the OpenTable __INITIAL_STATE__ JSON blob embedded in the page.
    Extracts website, phone, address, and email directly from structured data.
    Returns True if the JSON blob was found and parsed (even if some fields were empty).
    """
    soup = BeautifulSoup(html, "html.parser")
    script = soup.find("script", {"id": "primary-window-vars"})
    if not script:
        return False

    try:
        data = json.loads(script.string)
        rest = (
            data.get("windowVariables", {})
                .get("__INITIAL_STATE__", {})
                .get("restaurantProfile", {})
                .get("restaurant", {})
        )
        if not rest:
            return False

        # Website — skip PDF/download URLs (some chains register a menu PDF as their website)
        if not restaurant.get("website_url"):
            website = rest.get("website") or ""
            if website and is_restaurant_website(website):
                parsed = urllib.parse.urlparse(website)
                restaurant["website_url"] = urllib.parse.urlunparse(
                    (parsed.scheme, parsed.netloc, parsed.path, "", "", "")
                ).rstrip("/") + "/"

        # Contact info
        contact = rest.get("contactInformation", {})
        if not restaurant.get("phone") and contact.get("formattedPhoneNumber"):
            restaurant["phone"] = contact["formattedPhoneNumber"]
        if not restaurant.get("email") and contact.get("email"):
            restaurant["email"] = contact["email"]

        # Location count — try common field names OpenTable uses
        if not restaurant.get("location_count"):
            total = (
                rest.get("totalLocations")
                or rest.get("locationCount")
                or rest.get("numberOfLocations")
            )
            if total is not None:
                restaurant["location_count"] = int(total)

        # Address — build from structured fields
        if not restaurant.get("address"):
            addr = rest.get("address", {})
            parts = [
                addr.get("line1", ""),
                addr.get("line2", ""),
                addr.get("city", ""),
                addr.get("state", ""),
                addr.get("postCode", ""),
            ]
            address_str = ", ".join(p for p in parts if p)
            if address_str:
                restaurant["address"] = address_str

        return True

    except (json.JSONDecodeError, KeyError, TypeError) as exc:
        log.debug("JSON parse failed for %s: %s", restaurant.get("restaurant_name"), exc)
        return False


def enrich_from_opentable_bdata(restaurant: dict) -> dict:
    """
    Uses `bdata scrape` (Bright Data CLI) to fetch the OpenTable detail page,
    bypassing Akamai bot protection. Extracts website_url, phone, address, and email
    from the __INITIAL_STATE__ JSON blob embedded in the page.
    Install: sudo npm install -g @brightdata/cli && bdata login
    """
    opentable_url = restaurant.get("opentable_url", "")
    if not opentable_url:
        return restaurant

    try:
        html = None
        for attempt in range(1, 3):  # up to 2 attempts
            html = bdata_fetch_html(opentable_url, timeout=90)
            if html:
                break
            log.warning("bdata attempt %d failed for %s", attempt, opentable_url)
            time.sleep(2)

        if not html:
            return restaurant

        # Sanity check — Akamai block pages are tiny JS stubs
        if "sec-if-cpt-container" in html or len(html) < 5000:
            log.warning("Possible block page for %s (size=%d)", opentable_url, len(html))
            return restaurant

        # Primary: parse structured JSON from OpenTable's __INITIAL_STATE__ blob
        json_ok = _extract_from_ot_json(html, restaurant)

        if not json_ok:
            # Fallback: CSS/regex extraction (should rarely be needed)
            log.debug("JSON extraction failed for %s, falling back to HTML parsing", opentable_url)
            soup = BeautifulSoup(html, "html.parser")

            if not restaurant.get("phone"):
                tel = soup.select_one("a[href^='tel:']")
                if tel:
                    raw = tel.get("href", "").replace("tel:", "").strip()
                    if raw.startswith("+44"):
                        raw = "0" + re.sub(r"\s+", " ", raw[3:]).lstrip()
                    restaurant["phone"] = raw

            if not restaurant.get("address"):
                text = soup.get_text(" ", strip=True)
                postcode_match = ADDRESS_REGEX.search(text)
                if postcode_match:
                    preceding = text[max(0, postcode_match.start() - 200):postcode_match.start()]
                    house_nums = list(re.finditer(r'\d[\d\-]*\s+[A-Z]', preceding))
                    if house_nums:
                        addr_start = max(0, postcode_match.start() - 200) + house_nums[-1].start()
                    else:
                        addr_start = max(0, postcode_match.start() - 80)
                    restaurant["address"] = text[addr_start:postcode_match.end()].strip()

            if not restaurant.get("email"):
                emails = extract_emails_from_html(html)
                if emails:
                    restaurant["email"] = ", ".join(emails)

            if not restaurant.get("website_url"):
                soup2 = BeautifulSoup(html, "html.parser") if not json_ok else soup
                label = soup2.find("span", {"data-test": "restaurant-detail-title"},
                                   string=re.compile(r"^Website$", re.I))
                if label:
                    container = label.find_parent()
                    if container:
                        a = container.find("a", href=re.compile(r"^http"))
                        if a and is_restaurant_website(a["href"]):
                            parsed = urllib.parse.urlparse(a["href"])
                            restaurant["website_url"] = urllib.parse.urlunparse(
                                (parsed.scheme, parsed.netloc, parsed.path, "", "", "")
                            ).rstrip("/") + "/"

        # Derive lead_tier from location_count
        if not restaurant.get("lead_tier"):
            restaurant["lead_tier"] = get_lead_tier(restaurant.get("location_count"))

        log.info("Enriched: %s | phone=%s | address=%s | website=%s | email=%s | locations=%s | tier=%s",
                 restaurant.get("restaurant_name"),
                 restaurant.get("phone"), restaurant.get("address"),
                 restaurant.get("website_url"), restaurant.get("email"),
                 restaurant.get("location_count"), restaurant.get("lead_tier"))

    except FileNotFoundError:
        log.error("bdata CLI not found — install: sudo npm install -g @brightdata/cli && bdata login")
    except Exception as exc:
        log.warning("enrich_from_opentable_bdata failed for %s: %s", opentable_url, exc)

    return restaurant


# ── Google Sheets ─────────────────────────────────────────────────────────────

SHEET_COLUMNS = [
    "city",
    "location_count",
    "lead_tier",
    "location",
    "cuisine",
    "price_range",
    "scraped_at",
    "restaurant_name",
    "email",
    "phone",
    "address",
    "website_url",
    "instagram_url",
    "opentable_url",
    "rating",
    "review_count",
    "email_sent",
    "sent_at",
    "replied",
]


def get_sheet(tab_name: str):
    key_file = os.environ["GOOGLE_SHEETS_KEY_FILE"]
    sheet_id = os.environ["GOOGLE_SHEETS_ID"]
    creds = Credentials.from_service_account_file(key_file, scopes=GSHEETS_SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    try:
        sheet = spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=tab_name, rows=5000, cols=len(SHEET_COLUMNS))
        sheet.append_row(SHEET_COLUMNS)
        log.info("Created new tab: %s", tab_name)

    return sheet


def clear_sheet(sheet) -> None:
    """Clears all data and re-adds the header row."""
    sheet.clear()
    sheet.append_row(SHEET_COLUMNS)
    log.info("Sheet cleared (header re-added).")


def load_existing_urls(sheet) -> set[str]:
    """Returns a set of opentable_urls already in the sheet to avoid duplicates."""
    records = sheet.get_all_records()
    return {r["opentable_url"] for r in records if r.get("opentable_url")}


def load_existing_records(sheet) -> tuple[dict, list[str]]:
    """
    Returns ({opentable_url: {'row_idx': int, 'data': dict}}, headers).
    row_idx is 1-based (row 2 = first data row, since row 1 is header).
    Also returns the actual sheet headers so callers can use them for column lookups.
    """
    all_values = sheet.get_all_values()
    if len(all_values) < 2:
        return {}, []
    headers = all_values[0]
    records = {}
    for i, row in enumerate(all_values[1:], start=2):
        padded = row + [""] * max(0, len(headers) - len(row))
        row_dict = dict(zip(headers, padded))
        url = row_dict.get("opentable_url", "")
        if url:
            records[url] = {"row_idx": i, "data": row_dict}
    return records, headers


def update_row_fields(sheet, row_idx: int, updates: dict, sheet_headers: list[str] | None = None) -> None:
    """
    Updates specific cells in a sheet row by column name. Only writes non-empty values.
    Uses sheet_headers for column index lookup when provided — required when the sheet
    column order differs from SHEET_COLUMNS (e.g. manually added columns).
    Falls back to SHEET_COLUMNS if sheet_headers is not given.
    Retries each cell up to 3 times on network timeout.
    """
    col_source = sheet_headers if sheet_headers is not None else SHEET_COLUMNS
    for col_name, value in updates.items():
        if col_name in col_source and value:
            col_idx = col_source.index(col_name) + 1  # gspread is 1-indexed
            for attempt in range(1, 4):
                try:
                    sheet.update_cell(row_idx, col_idx, value)
                    break
                except Exception as exc:
                    if attempt < 3:
                        log.warning("  Sheet write failed (attempt %d/3): %s — retrying in 5s", attempt, exc)
                        time.sleep(5)
                    else:
                        log.error("  Sheet write failed after 3 attempts for %s: %s", col_name, exc)


def write_contacts_to_sheet(sheet, contacts: list[dict], existing_urls: set[str]) -> int:
    """Appends new contacts to the sheet. Deduplicates by opentable_url. Returns count of rows written."""
    rows_written = 0
    now = datetime.utcnow().strftime("%Y-%m-%d")

    for contact in contacts:
        opentable_url = contact.get("opentable_url", "")
        if opentable_url in existing_urls:
            log.info("Skipping duplicate: %s", contact.get("restaurant_name"))
            continue

        row = [contact.get(col, "") or "" for col in SHEET_COLUMNS]
        row[SHEET_COLUMNS.index("scraped_at")] = now

        sheet.append_row(row)
        existing_urls.add(opentable_url)
        rows_written += 1

    return rows_written


# ── Main ──────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(description="Scrape OpenTable UK restaurants and extract contact emails.")
    parser.add_argument(
        "--cities",
        nargs="+",
        default=DEFAULT_CITIES,
        help="Cities to scrape (default: %(default)s)",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Max restaurants per city (useful for testing)",
    )
    parser.add_argument(
        "--no-gsheets",
        action="store_true",
        help="Dry run — print results to console without writing to Google Sheets",
    )
    parser.add_argument(
        "--clear",
        action="store_true",
        help="Clear all existing rows in the sheet before writing (keeps header)",
    )
    parser.add_argument(
        "--fill-missing",
        action="store_true",
        help="Scan existing sheet rows and fill in any missing email/phone/address/website_url",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    log.info("Cities: %s", args.cities)
    log.info("Limit per city: %s", args.limit or "none")
    log.info("Google Sheets write: %s", not args.no_gsheets)

    # ── Fill-missing mode ─────────────────────────────────────────────────────
    if args.fill_missing:
        sheet = get_sheet(CONTACTS_TAB)
        records, sheet_headers = load_existing_records(sheet)
        FILL_FIELDS = ["location_count", "lead_tier", "email", "phone", "address", "website_url", "instagram_url"]
        incomplete = {
            url: rec for url, rec in records.items()
            if any(not rec["data"].get(f) for f in FILL_FIELDS)
        }
        log.info("Rows to fill: %d / %d total", len(incomplete), len(records))

        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            )
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                locale="en-GB",
                viewport={"width": 1280, "height": 900},
            )
            Stealth(navigator_languages_override=("en-GB", "en")).apply_stealth_sync(context)
            page = context.new_page()

            for i, (url, rec) in enumerate(incomplete.items(), start=1):
                name = rec["data"].get("restaurant_name", url)
                log.info("[fill %d/%d] %s", i, len(incomplete), name)

                # Build a restaurant dict with empty strings normalised to None
                restaurant = {k: (v or None) for k, v in rec["data"].items()}
                restaurant["opentable_url"] = url

                # Derive lead_tier from existing location_count if already present
                if restaurant.get("location_count") and not restaurant.get("lead_tier"):
                    restaurant["lead_tier"] = get_lead_tier(restaurant["location_count"])

                # Step 2: re-fetch OpenTable detail if any structural field is missing
                if (not restaurant.get("phone") or not restaurant.get("website_url")
                        or not restaurant.get("address") or not restaurant.get("location_count")):
                    restaurant = enrich_from_opentable_bdata(restaurant)
                    time.sleep(CRAWL_DELAY)

                # Step 3: visit restaurant website if email or instagram_url still missing
                if (not restaurant.get("email") or not restaurant.get("instagram_url")) and restaurant.get("website_url"):
                    details = fetch_contact_details(page, restaurant["website_url"])
                    restaurant.update(details)

                # Write back only fields that were empty and are now filled
                updates = {
                    f: restaurant[f]
                    for f in FILL_FIELDS
                    if restaurant.get(f) and not rec["data"].get(f)
                }
                if updates:
                    update_row_fields(sheet, rec["row_idx"], updates, sheet_headers)
                    log.info("  → filled: %s", list(updates.keys()))
                else:
                    log.info("  → nothing new found")

            browser.close()

        log.info("Fill-missing done.")
        return

    all_contacts = []

    # Open sheet once upfront so we can write after each city
    if not args.no_gsheets:
        sheet = get_sheet(CONTACTS_TAB)
        if args.clear:
            clear_sheet(sheet)
        existing_urls = load_existing_urls(sheet)
        log.info("Existing contacts in sheet: %d", len(existing_urls))
    else:
        sheet = None
        existing_urls = set()

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        context = browser.new_context(
            user_agent=HEADERS["User-Agent"],
            locale="en-GB",
            viewport={"width": 1280, "height": 900},
            extra_http_headers={
                "Accept-Language": "en-GB,en;q=0.9",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
                "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"macOS"',
            },
        )
        Stealth(navigator_languages_override=("en-GB", "en")).apply_stealth_sync(context)
        page = context.new_page()

        for city in args.cities:
            # ── Step 1: Scrape listings ───────────────────────────────────────
            t1_start = time.time()
            log.info("[%s] ── Step 1 start: scraping OpenTable listings ──", city)
            listings = scrape_opentable_city(page, city, limit=args.limit)
            t1_end = time.time()
            log.info("[%s] ── Step 1 done: %d listings | duration: %.0fs ──",
                     city, len(listings), t1_end - t1_start)

            # Filter out restaurants already in the sheet before hitting bdata
            new_listings = [l for l in listings if l.get("opentable_url") not in existing_urls]
            skipped = len(listings) - len(new_listings)
            if skipped:
                log.info("[%s] Skipping %d already-scraped restaurants", city, skipped)

            # ── Step 2: OpenTable detail pages → phone, address, website_url ─
            t2_start = time.time()
            log.info("[%s] ── Step 2 start: enriching from OpenTable detail pages ──", city)
            enriched = []
            for i, listing in enumerate(new_listings):
                log.info("[%s] Step 2 [%d/%d]: %s", city, i + 1, len(new_listings), listing.get("restaurant_name"))
                enriched.append(enrich_from_opentable_bdata(listing))
                time.sleep(CRAWL_DELAY)
            t2_end = time.time()
            with_website = sum(1 for r in enriched if r.get("website_url"))
            log.info("[%s] ── Step 2 done: %d/%d with website | duration: %.0fs (avg %.0fs/restaurant) ──",
                     city, with_website, len(enriched), t2_end - t2_start,
                     (t2_end - t2_start) / max(len(enriched), 1))

            # ── Step 3: Restaurant website → email ───────────────────────────
            t3_start = time.time()
            log.info("[%s] ── Step 3 start: scraping emails from restaurant websites ──", city)
            for i, restaurant in enumerate(enriched):
                website = restaurant.get("website_url")
                already_has_email = bool(restaurant.get("email"))
                if website and not already_has_email:
                    log.info("[%s] Step 3 [%d/%d]: %s", city, i + 1, len(enriched), website)
                    details = fetch_contact_details(page, website)
                    restaurant.update(details)
                elif already_has_email:
                    log.info("[%s] Step 3 [%d/%d]: skipped (email from OT: %s)",
                             city, i + 1, len(enriched), restaurant["email"])
                time.sleep(CRAWL_DELAY)
            t3_end = time.time()
            found_emails = sum(1 for r in enriched if r.get("email"))
            log.info("[%s] ── Step 3 done: %d/%d with email | duration: %.0fs (avg %.0fs/restaurant) ──",
                     city, found_emails, len(enriched), t3_end - t3_start,
                     (t3_end - t3_start) / max(len(enriched), 1))

            log.info("[%s] ── City total duration: %.0fs ──", city, t3_end - t1_start)
            all_contacts.extend(enriched)

            # ── Step 4: Write city results to sheet immediately ───────────────
            if not args.no_gsheets:
                written = write_contacts_to_sheet(sheet, enriched, existing_urls)
                log.info("[%s] ── Step 4 done: %d new rows written to sheet ──", city, written)
            else:
                log.info("[%s] --no-gsheets: skipping sheet write", city)

        browser.close()

    total_with_email = sum(1 for c in all_contacts if c.get("email"))
    log.info("Done. Total scraped: %d | With email: %d", len(all_contacts), total_with_email)

    if args.no_gsheets:
        log.info("--no-gsheets flag set. Printing sample (first 5):")
        for c in all_contacts[:5]:
            print(json.dumps(c, indent=2))


if __name__ == "__main__":
    main()
