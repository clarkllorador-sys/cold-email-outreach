"""
send_emails.py
Reads business contacts from Google Sheets (contacts tab),
composes a personalised cold email with a Loom video thumbnail, sends via Gmail API,
and marks sent rows with email_sent=TRUE and sent_at timestamp.

Usage:
    python send_emails.py                    # send up to daily limit (default 10)
    python send_emails.py --limit 5          # override daily limit
    python send_emails.py --dry-run          # print emails, don't send
    python send_emails.py --only-with-email  # skip rows missing email (default: also skip)

Requirements (.env):
    GOOGLE_SHEETS_KEY_FILE   path to service account JSON
    GOOGLE_SHEETS_ID         spreadsheet ID
    GMAIL_FROM               sender address (e.g. you@yourdomain.com)
    LOOM_VIDEO_URL           Loom share URL (e.g. https://www.loom.com/share/abc123)
    LOOM_THUMBNAIL_URL       direct URL to thumbnail image (from Loom embed metadata)

Gmail API setup (one-time):
    1. Enable Gmail API in Google Cloud Console
    2. Create OAuth 2.0 credentials → download as gmail_credentials.json
    3. On first run, browser opens for auth → token saved to gmail_token.json
    OR: use a service account with domain-wide delegation (G Suite only)
"""

import argparse
import base64
import json
import logging
import os
import time
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────

DAILY_LIMIT_DEFAULT = 10          # emails/day — safe for cold outreach
SEND_DELAY = 3.0                  # seconds between sends (avoid rate limits)

CONTACTS_TAB = "contacts"

GSHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# Gmail OAuth token/credentials paths
GMAIL_CREDENTIALS_FILE = os.getenv("GMAIL_CREDENTIALS_FILE", "credentials/oath_client.json")
GMAIL_TOKEN_FILE = os.getenv("GMAIL_TOKEN_FILE", "credentials/gmail_token.json")

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Email template ────────────────────────────────────────────────────────────

EMAIL_SUBJECT = "Quick video for {restaurant_name}"

BOOKING_URL = "https://cal.example.com/free-consultation"

EMAIL_BODY_HTML = """\
<html>
<body style="font-family: Arial, sans-serif; font-size: 15px; color: #222; line-height: 1.6; max-width: 600px;">

<p>Hey {restaurant_name} team,</p>

<p>Hope you don't mind the cold reach-out ;')</p>

<p>Came across you on OpenTable and thought it was worth getting in touch. I put together a short video mentioning some of the results we've managed to accomplish.</p>

<p>It's about getting more bookings and visibility online. We work on a results-first basis, so there's really nothing to lose by having a look :)</p>

<p>
  <a href="{loom_url}" target="_blank" style="display:inline-block; text-decoration:none;">
    <img src="{thumbnail_url}"
         alt="Watch Video"
         width="480"
         style="border-radius:8px; border:2px solid #ddd; display:block;" />
  </a>
</p>

<p><a href="{booking_url}" target="_blank" style="color:#222;">Book a 20 minute call here</a></p>

<p>Would love to know what you think!</p>

<p>
  Alex<br/>
  Founder
</p>

</body>
</html>
"""

EMAIL_BODY_PLAIN = """\
Hey {restaurant_name} team,

Hope you don't mind the cold reach-out ;')

Came across you on OpenTable and thought it was worth getting in touch. I put together a short video mentioning some of the results we've managed to accomplish.

It's about getting more bookings and visibility online. We work on a results-first basis, so there's really nothing to lose by having a look :)

Watch video: {loom_url}

Book a 20 minute call here: {booking_url}

Would love to know what you think!

Alex
Founder
"""

# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_sheet_client():
    key_file = os.environ["GOOGLE_SHEETS_KEY_FILE"]
    creds = SACredentials.from_service_account_file(key_file, scopes=GSHEETS_SCOPES)
    return gspread.authorize(creds)


def load_contacts(sheet_id: str) -> tuple[gspread.Worksheet, list[str], list[dict]]:
    """Returns (worksheet, header_row, list_of_row_dicts)."""
    client = get_sheet_client()
    sheet = client.open_by_key(sheet_id)
    ws = sheet.worksheet(CONTACTS_TAB)
    all_rows = ws.get_all_values()
    if not all_rows:
        return ws, [], []
    headers = all_rows[0]
    rows = [dict(zip(headers, row)) for row in all_rows[1:]]
    return ws, headers, rows


def mark_sent(ws: gspread.Worksheet, headers: list[str], row_index_1based: int):
    """Updates email_sent=TRUE and sent_at=now for the given 1-based data row (row 2 = index 1)."""
    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    sheet_row = row_index_1based + 1  # +1 for header row

    col_sent = headers.index("email_sent") + 1
    col_at   = headers.index("sent_at") + 1

    ws.update_cell(sheet_row, col_sent, "TRUE")
    ws.update_cell(sheet_row, col_at, now_str)
    log.info("  ✓ Marked row %d as sent (%s)", sheet_row, now_str)

# ── Gmail API ─────────────────────────────────────────────────────────────────

def get_gmail_service():
    """
    Returns an authenticated Gmail API service.
    Uses OAuth 2.0 with local token cache (gmail_token.json).
    On first run: opens browser for consent. Subsequent runs: uses cached token.
    """
    creds = None

    if Path(GMAIL_TOKEN_FILE).exists():
        creds = OAuthCredentials.from_authorized_user_file(GMAIL_TOKEN_FILE, GMAIL_SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not Path(GMAIL_CREDENTIALS_FILE).exists():
                raise FileNotFoundError(
                    f"Gmail credentials not found at {GMAIL_CREDENTIALS_FILE}. "
                    "Download OAuth 2.0 credentials from Google Cloud Console."
                )
            flow = InstalledAppFlow.from_client_secrets_file(GMAIL_CREDENTIALS_FILE, GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        Path(GMAIL_TOKEN_FILE).write_text(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def build_message(
    from_addr: str,
    to_addr: str,
    subject: str,
    body_html: str,
    body_plain: str,
) -> dict:
    """Builds a base64url-encoded Gmail API message dict."""
    msg = MIMEMultipart("alternative")
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg["Subject"] = subject

    msg.attach(MIMEText(body_plain, "plain"))
    msg.attach(MIMEText(body_html, "html"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}


def send_message(service, message: dict) -> str:
    """Sends a Gmail message. Returns message ID."""
    result = service.users().messages().send(userId="me", body=message).execute()
    return result["id"]

# ── Core logic ────────────────────────────────────────────────────────────────

def pick_first_email(email_field: str) -> str | None:
    """
    email column may contain comma-separated addresses.
    Returns the first valid-looking one.
    """
    if not email_field or not email_field.strip():
        return None
    candidates = [e.strip() for e in email_field.split(",")]
    for candidate in candidates:
        if "@" in candidate and "." in candidate.split("@")[-1]:
            return candidate
    return None


def run(limit: int, dry_run: bool, only_to: str = ""):
    sheet_id = os.environ["GOOGLE_SHEETS_ID"]
    from_addr = os.environ["GMAIL_FROM"]
    if only_to:
        log.info("--only-to filter active: will only send to %s", only_to)

    log.info("Loading contacts from sheet %s ...", sheet_id)
    ws, headers, rows = load_contacts(sheet_id)

    if not rows:
        log.info("No rows found in sheet.")
        return

    # Validate required columns exist
    required_cols = {"restaurant_name", "email", "email_sent"}
    missing = required_cols - set(headers)
    if missing:
        raise ValueError(f"Sheet is missing columns: {missing}")

    # Filter: unsent rows that have an email AND a video URL ready
    pending = []
    for i, row in enumerate(rows):
        if row.get("email_sent", "").strip().upper() == "TRUE":
            continue
        email = pick_first_email(row.get("email", ""))
        if not email:
            log.debug("Row %d (%s) — no email, skipping", i + 2, row.get("restaurant_name"))
            continue
        # Safety filter: --only-to restricts sends to one address
        if only_to and email.lower() != only_to.lower():
            log.debug("Row %d (%s) — skipping (--only-to filter)", i + 2, row.get("restaurant_name"))
            continue
        video_url = row.get("drive_video_url", "").strip()
        thumbnail_url = row.get("drive_thumbnail_url", "").strip()
        if not video_url:
            log.debug("Row %d (%s) — no video yet (run generate_videos.py first), skipping",
                      i + 2, row.get("restaurant_name"))
            continue
        pending.append((i, row, email, video_url, thumbnail_url))

    log.info("%d unsent rows with email + video found (limit=%d)", len(pending), limit)

    if not pending:
        log.info("Nothing to send. (Have you run generate_videos.py yet?)")
        return

    # Init Gmail only if actually sending
    service = None
    if not dry_run:
        log.info("Authenticating Gmail ...")
        service = get_gmail_service()

    sent = 0
    for i, row, to_email, video_url, thumbnail_url in pending:
        if sent >= limit:
            log.info("Daily limit of %d reached. Stopping.", limit)
            break

        restaurant_name = row.get("restaurant_name", "").strip() or "the team"

        subject = EMAIL_SUBJECT.format(restaurant_name=restaurant_name)
        body_html = EMAIL_BODY_HTML.format(
            restaurant_name=restaurant_name,
            loom_url=video_url,
            thumbnail_url=thumbnail_url,
            booking_url=BOOKING_URL,
        )
        body_plain = EMAIL_BODY_PLAIN.format(
            restaurant_name=restaurant_name,
            loom_url=video_url,
            booking_url=BOOKING_URL,
        )

        log.info("[%d/%d] %s → %s  (video: %s)", sent + 1, min(limit, len(pending)), restaurant_name, to_email, video_url[:60])

        if dry_run:
            print(f"\n{'='*60}")
            print(f"TO:      {to_email}")
            print(f"SUBJECT: {subject}")
            print(f"BODY:\n{body_plain}")
            sent += 1
            continue

        try:
            msg = build_message(from_addr, to_email, subject, body_html, body_plain)
            msg_id = send_message(service, msg)
            log.info("  Sent (Gmail ID: %s)", msg_id)
            mark_sent(ws, headers, i + 1)  # i is 0-based in rows[], +1 for 1-based
            sent += 1
            time.sleep(SEND_DELAY)
        except HttpError as exc:
            log.error("  Gmail API error for %s: %s", to_email, exc)
        except Exception as exc:
            log.error("  Unexpected error for %s: %s", to_email, exc)

    log.info("Done. %d email(s) sent.", sent)

# ── CLI ───────────────────────────────────────────────────────────────────────

def youtube_thumbnail(url: str) -> str:
    """Derives maxresdefault thumbnail URL from a YouTube watch URL."""
    import re
    m = re.search(r"(?:v=|youtu\.be/)([A-Za-z0-9_-]{11})", url)
    if m:
        return f"https://img.youtube.com/vi/{m.group(1)}/maxresdefault.jpg"
    raise ValueError(f"Could not extract YouTube video ID from: {url}")


def run_test(to_addr: str, video_url: str, from_addr: str):
    """Send a single test email — no sheet interaction."""
    thumbnail_url = youtube_thumbnail(video_url) if "youtube.com" in video_url or "youtu.be" in video_url \
        else os.environ.get("LOOM_THUMBNAIL_URL", "")

    restaurant_name = "the team"
    subject = EMAIL_SUBJECT.format(restaurant_name="Test Co")
    body_html = EMAIL_BODY_HTML.format(
        restaurant_name=restaurant_name,
        loom_url=video_url,
        thumbnail_url=thumbnail_url,
        booking_url=BOOKING_URL,
    )
    body_plain = EMAIL_BODY_PLAIN.format(
        restaurant_name=restaurant_name,
        loom_url=video_url,
        booking_url=BOOKING_URL,
    )

    log.info("Test email: %s → %s", from_addr, to_addr)
    log.info("Thumbnail: %s", thumbnail_url)
    log.info("Authenticating Gmail ...")
    service = get_gmail_service()
    msg = build_message(from_addr, to_addr, subject, body_html, body_plain)
    msg_id = send_message(service, msg)
    log.info("✓ Sent (Gmail ID: %s)", msg_id)


def main():
    parser = argparse.ArgumentParser(description="Cold email sender (Gmail API)")
    parser.add_argument(
        "--limit",
        type=int,
        default=DAILY_LIMIT_DEFAULT,
        help=f"Max emails to send this run (default: {DAILY_LIMIT_DEFAULT})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print emails to stdout without sending or updating the sheet",
    )
    parser.add_argument(
        "--test",
        metavar="TO_EMAIL",
        help="Send a single test email to this address (bypasses sheet)",
    )
    parser.add_argument(
        "--video-url",
        default=os.getenv("LOOM_VIDEO_URL", ""),
        help="Video URL to embed (YouTube or Loom). Overrides LOOM_VIDEO_URL env var.",
    )
    parser.add_argument(
        "--from",
        dest="from_addr",
        default=os.getenv("GMAIL_FROM", ""),
        help="Sender address. Overrides GMAIL_FROM env var.",
    )
    parser.add_argument(
        "--only-to",
        metavar="EMAIL",
        default="",
        help="Safety filter: only send to this address (ignores all other rows).",
    )
    args = parser.parse_args()

    if args.test:
        if not args.video_url:
            parser.error("--video-url is required for --test (or set LOOM_VIDEO_URL in .env)")
        if not args.from_addr:
            parser.error("--from is required for --test (or set GMAIL_FROM in .env)")
        run_test(to_addr=args.test, video_url=args.video_url, from_addr=args.from_addr)
        return

    if args.dry_run:
        log.info("DRY RUN — no emails will be sent.")

    run(limit=args.limit, dry_run=args.dry_run, only_to=args.only_to)


if __name__ == "__main__":
    main()
