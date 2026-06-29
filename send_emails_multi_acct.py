"""
send_emails_multi_acct.py
Multi-account cold email sender via SMTP (SiteGround).
Reads restaurant contacts from Google Sheets (contacts tab),
composes a personalised cold email with a video thumbnail, sends via SMTP,
and marks sent rows with email_sent=TRUE and sent_at timestamp.

Usage:
    python send_emails_multi_acct.py                        # send up to daily limit (default 15)
    python send_emails_multi_acct.py --limit 30             # override daily limit
    python send_emails_multi_acct.py --dry-run              # print emails, don't send
    python send_emails_multi_acct.py --only-to EMAIL        # safety filter: send only to this address
    python send_emails_multi_acct.py --send-mode first      # first email per restaurant only (default)
    python send_emails_multi_acct.py --send-mode all        # all emails found per restaurant
    python send_emails_multi_acct.py --send-mode max2       # max 2 emails per restaurant
    python send_emails_multi_acct.py --test TO_EMAIL        # send one test email, no sheet interaction

Requirements (.env):
    GOOGLE_SHEETS_KEY_FILE      path to service account JSON
    GOOGLE_SHEETS_ID            spreadsheet ID

    SENDER_1_FROM=you@sender1.com
    SENDER_1_PASSWORD=...
    SENDER_1_HOST=mail.sender1.com

    SENDER_2_FROM=you@sender2.com
    SENDER_2_PASSWORD=...
    SENDER_2_HOST=mail.sender2.com

    SENDER_3_FROM=you@sender3.com
    SENDER_3_PASSWORD=...
    SENDER_3_HOST=mail.sender3.com

    SMTP_PORT=465               (SSL — default)
"""

import argparse
import logging
import os
import smtplib
import ssl
import time
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials as SACredentials

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────

DAILY_LIMIT_DEFAULT = 15          # total emails/day across all senders (warm-up phase)
SEND_DELAY = 3.0                  # seconds between sends
CONTACTS_TAB = "contacts"
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))

GSHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Email template ────────────────────────────────────────────────────────────

EMAIL_SUBJECT = "I made this for you, {restaurant_name}"

BOOKING_URL = "https://cal.example.com/free-consultation"

EMAIL_BODY_HTML = """\
<html>
<body style="font-family: Arial, sans-serif; font-size: 15px; color: #222; line-height: 1.6; max-width: 600px;">

<p>Hey {restaurant_name} team,</p>

<p>Hope this isn't too out of the blue ;')</p>

<p>I came across you on OpenTable and put together a short personal video for you</p>

<p>It's only a minute, worth a watch :)</p>

<p>
  <strong>{restaurant_name} — watch here</strong><br/>
  <a href="{video_url}" target="_blank" style="display:inline-block; text-decoration:none;">
    <img src="{thumbnail_url}"
         alt="Watch Video"
         width="480"
         style="border-radius:8px; border:2px solid #ddd; display:block; margin-top:8px;" />
  </a>
</p>

<p><a href="{booking_url}" target="_blank" style="color:#1a73e8; text-decoration:underline; font-weight:bold;">Grab a free 20-minute chat</a></p>

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

Hope this isn't too out of the blue ;')

I came across you on OpenTable and put together a short personal video for you

It's only a minute, worth a watch :)

{restaurant_name} — watch here: {video_url}

Grab a free 20-minute chat: {booking_url}

Would love to know what you think!

Alex
Founder
"""

# ── Sender config ─────────────────────────────────────────────────────────────

def load_senders() -> list[dict]:
    """
    Loads up to 3 sender configs from .env.
    Returns list of dicts: {from, password, host}
    Raises if no senders configured.
    """
    senders = []
    for n in range(1, 4):
        from_addr = os.getenv(f"SENDER_{n}_FROM", "").strip()
        password  = os.getenv(f"SENDER_{n}_PASSWORD", "").strip()
        host      = os.getenv(f"SENDER_{n}_HOST", "").strip()
        if from_addr and password and host:
            senders.append({"from": from_addr, "password": password, "host": host})
        elif from_addr:
            log.warning("SENDER_%d_FROM set but missing PASSWORD or HOST — skipping", n)
    if not senders:
        raise ValueError(
            "No sender accounts configured. Add SENDER_1_FROM, SENDER_1_PASSWORD, "
            "SENDER_1_HOST (and optionally SENDER_2_*, SENDER_3_*) to .env"
        )
    return senders

# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_sheet_client():
    key_file = os.environ["GOOGLE_SHEETS_KEY_FILE"]
    creds = SACredentials.from_service_account_file(key_file, scopes=GSHEETS_SCOPES)
    return gspread.authorize(creds)


def load_contacts(sheet_id: str) -> tuple[gspread.Worksheet, list[str], list[dict]]:
    client = get_sheet_client()
    sheet = client.open_by_key(sheet_id)
    ws = sheet.worksheet(CONTACTS_TAB)
    all_rows = ws.get_all_values()
    if not all_rows:
        return ws, [], []
    headers = all_rows[0]
    rows = [dict(zip(headers, row)) for row in all_rows[1:]]
    return ws, headers, rows


def mark_sent(ws: gspread.Worksheet, headers: list[str], row_index_1based: int, from_addr: str):
    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    sheet_row = row_index_1based + 1  # +1 for header row

    col_sent = headers.index("email_sent") + 1
    col_at   = headers.index("sent_at") + 1

    ws.update_cell(sheet_row, col_sent, "TRUE")
    ws.update_cell(sheet_row, col_at, now_str)
    log.info("  ✓ Marked row %d as sent via %s (%s)", sheet_row, from_addr, now_str)

# ── SMTP sending ──────────────────────────────────────────────────────────────

def build_mime_message(
    from_addr: str,
    to_addr: str,
    subject: str,
    body_html: str,
    body_plain: str,
) -> MIMEMultipart:
    msg = MIMEMultipart("alternative")
    msg["From"] = f"Alex <{from_addr}>"
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body_plain, "plain"))
    msg.attach(MIMEText(body_html, "html"))
    return msg


def send_via_smtp(sender: dict, to_addr: str, msg: MIMEMultipart):
    """Send one email via SMTP SSL (port 465)."""
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(sender["host"], SMTP_PORT, context=context) as server:
        server.login(sender["from"], sender["password"])
        server.sendmail(sender["from"], to_addr, msg.as_string())

# ── Email address helpers ─────────────────────────────────────────────────────

def pick_emails(email_field: str, mode: str) -> list[str]:
    """
    Returns a list of email addresses to send to based on send_mode:
      first — first valid address only
      all   — all valid addresses
      max2  — up to 2 valid addresses
    """
    if not email_field or not email_field.strip():
        return []
    candidates = [
        e.strip() for e in email_field.split(",")
        if "@" in e and "." in e.split("@")[-1]
    ]
    if not candidates:
        return []
    if mode == "first":
        return [candidates[0]]
    elif mode == "max2":
        return candidates[:2]
    else:  # all
        return candidates

# ── Core logic ────────────────────────────────────────────────────────────────

def run(limit: int, dry_run: bool, send_mode: str, only_to: str = ""):
    sheet_id = os.environ["GOOGLE_SHEETS_ID"]
    senders = load_senders()

    log.info("Loaded %d sender account(s): %s", len(senders), [s["from"] for s in senders])
    log.info("Row limit: %d restaurants", limit)

    only_to_list = [e.strip().lower() for e in only_to.split(",") if e.strip()] if only_to else []
    if only_to_list:
        log.info("--only-to filter active: will only send to %s", only_to_list)

    log.info("Loading contacts from sheet %s ...", sheet_id)
    ws, headers, rows = load_contacts(sheet_id)

    if not rows:
        log.info("No rows found in sheet.")
        return

    required_cols = {"restaurant_name", "email", "email_sent"}
    missing = required_cols - set(headers)
    if missing:
        raise ValueError(f"Sheet is missing columns: {missing}")

    # Build pending list: (row_index_0based, row, [emails_to_send], video_url, thumbnail_url)
    pending = []
    for i, row in enumerate(rows):
        if row.get("email_sent", "").strip().upper() == "TRUE":
            continue
        emails = pick_emails(row.get("email", ""), send_mode)
        if not emails:
            log.debug("Row %d (%s) — no email, skipping", i + 2, row.get("restaurant_name"))
            continue
        if only_to_list:
            emails = only_to_list  # replace recipients — used for test sends
        video_url     = row.get("drive_video_url", "").strip()
        thumbnail_url = row.get("drive_thumbnail_url", "").strip()
        if not video_url:
            log.debug("Row %d (%s) — no video yet, skipping", i + 2, row.get("restaurant_name"))
            continue
        pending.append((i, row, emails, video_url, thumbnail_url))

    log.info("%d unsent rows with email + video found (limit=%d, mode=%s)", len(pending), limit, send_mode)

    if not pending:
        log.info("Nothing to send.")
        return

    # Assign rows to senders: round-robin across senders
    sender_counts = [0] * len(senders)
    total_sent = 0
    rows_processed = 0

    for i, row, emails, video_url, thumbnail_url in pending:
        if rows_processed >= limit:
            log.info("Row limit of %d reached. Stopping.", limit)
            break

        # Round-robin sender selection
        sender_idx = rows_processed % len(senders)
        sender = senders[sender_idx]

        restaurant_name = row.get("restaurant_name", "").strip() or "the team"

        for to_email in emails:

            subject = EMAIL_SUBJECT.format(restaurant_name=restaurant_name)
            body_html = EMAIL_BODY_HTML.format(
                restaurant_name=restaurant_name,
                video_url=video_url,
                thumbnail_url=thumbnail_url,
                booking_url=BOOKING_URL,
            )
            body_plain = EMAIL_BODY_PLAIN.format(
                restaurant_name=restaurant_name,
                video_url=video_url,
                booking_url=BOOKING_URL,
            )

            log.info(
                "[%d] %s → %s  (from: %s)",
                total_sent + 1, restaurant_name, to_email, sender["from"]
            )

            if dry_run:
                print(f"\n{'='*60}")
                print(f"FROM:    {sender['from']}")
                print(f"TO:      {to_email}")
                print(f"SUBJECT: {subject}")
                print(f"BODY:\n{body_plain}")
                total_sent += 1
                continue

            try:
                msg = build_mime_message(sender["from"], to_email, subject, body_html, body_plain)
                send_via_smtp(sender, to_email, msg)
                log.info("  ✓ Sent")
                total_sent += 1
                time.sleep(SEND_DELAY)
            except Exception as exc:
                log.error("  ✗ Failed to send to %s: %s", to_email, exc)
                continue

        rows_processed += 1
        sender_counts[sender_idx] += 1

        # Mark row as sent after processing all emails for this restaurant
        if not dry_run:
            try:
                mark_sent(ws, headers, i + 1, sender["from"])
            except Exception as exc:
                log.error("  ✗ Failed to mark row as sent: %s", exc)

    log.info("Done. %d restaurant(s) processed, %d email(s) sent. Sender breakdown: %s",
             rows_processed, total_sent,
             {senders[idx]["from"]: sender_counts[idx] for idx in range(len(senders))})

# ── Test mode ─────────────────────────────────────────────────────────────────

def run_test(to_addr: str, sender_num: int, video_url: str = ""):
    """Send a single test email using the specified sender (1-indexed). No sheet interaction."""
    senders = load_senders()
    idx = sender_num - 1
    if idx < 0 or idx >= len(senders):
        raise ValueError(f"--sender {sender_num} is out of range — only {len(senders)} sender(s) configured")

    sender = senders[idx]
    restaurant_name = "Test Co"
    video_url = video_url or "https://www.google.com"
    thumbnail_url = ""

    subject = EMAIL_SUBJECT.format(restaurant_name=restaurant_name)
    body_html = EMAIL_BODY_HTML.format(
        restaurant_name=restaurant_name,
        video_url=video_url,
        thumbnail_url=thumbnail_url or "https://via.placeholder.com/480x270?text=Video+Thumbnail",
        booking_url=BOOKING_URL,
    )
    body_plain = EMAIL_BODY_PLAIN.format(
        restaurant_name=restaurant_name,
        video_url=video_url,
        booking_url=BOOKING_URL,
    )

    log.info("Test email: %s → %s (SMTP: %s:%s)", sender["from"], to_addr, sender["host"], SMTP_PORT)
    msg = build_mime_message(sender["from"], to_addr, subject, body_html, body_plain)
    send_via_smtp(sender, to_addr, msg)
    log.info("✓ Sent successfully from %s", sender["from"])

# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Multi-account cold email sender (SMTP)")
    parser.add_argument(
        "--limit", type=int, default=DAILY_LIMIT_DEFAULT,
        help=f"Total emails to send this run across all senders (default: {DAILY_LIMIT_DEFAULT})",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Print emails to stdout without sending or updating the sheet",
    )
    parser.add_argument(
        "--send-mode", choices=["first", "all", "max2"], default="first",
        help="How many emails to send per restaurant: first (default), all, or max2",
    )
    parser.add_argument(
        "--only-to", metavar="EMAIL", default="",
        help="Safety filter: only send to this address (ignores all other rows)",
    )
    parser.add_argument(
        "--test", metavar="TO_EMAIL",
        help="Send a single test email to this address (bypasses sheet)",
    )
    parser.add_argument(
        "--sender", type=int, default=1, metavar="N",
        help="Which sender to use for --test (1, 2, or 3 — default: 1)",
    )
    parser.add_argument(
        "--video-url", default="",
        help="Real video URL to use in --test mode (optional)",
    )
    args = parser.parse_args()

    if args.test:
        run_test(to_addr=args.test, sender_num=args.sender, video_url=args.video_url)
        return

    if args.dry_run:
        log.info("DRY RUN — no emails will be sent.")

    run(
        limit=args.limit,
        dry_run=args.dry_run,
        send_mode=args.send_mode,
        only_to=args.only_to,
    )


if __name__ == "__main__":
    main()
