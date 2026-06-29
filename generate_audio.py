"""
generate_audio.py
For each sheet row without audio_verified=TRUE:
  1. Call ElevenLabs TTS API → restaurant name in host's cloned voice → MP3
  2. Run Whisper STT on the MP3 → transcription
  3. Fuzzy-match transcription vs restaurant_name → auto-verify
  4. Upload MP3 to Google Drive → drive_audio_url
  5. Write drive_audio_url, audio_transcription, audio_verified to sheet

Run BEFORE generate_videos.py — generate_videos.py will skip rows
where audio_verified != TRUE.

Usage:
    python generate_audio.py                        # all unverified rows
    python generate_audio.py --restaurant "Dishoom" # single restaurant
    python generate_audio.py --limit 5              # cap at 5
    python generate_audio.py --dry-run              # print without generating

Requirements (.env):
    ELEVENLABS_API_KEY        ElevenLabs API key
    ELEVENLABS_VOICE_ID       Cloned voice ID from ElevenLabs Voices page
    GOOGLE_SHEETS_KEY_FILE    service account JSON
    GOOGLE_SHEETS_ID          spreadsheet ID
    GOOGLE_DRIVE_FOLDER_ID    Drive folder for audio files

Sheet columns added:
    drive_audio_url       Google Drive shareable link to MP3
    audio_transcription   What Whisper heard
    audio_verified        TRUE/FALSE — auto-set based on STT match
"""

import argparse
import logging
import os
import re
import subprocess
import tempfile
import time
from pathlib import Path

import requests
import gspread
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────

CONTACTS_TAB = "contacts"

ELEVENLABS_TTS_URL = "https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"
ELEVENLABS_MODEL = "eleven_multilingual_v2"  # matches UI settings that sounded best

WHISPER_MODEL = "base"  # free, local, fast enough for 1-2 second clips

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

# ── ElevenLabs TTS ────────────────────────────────────────────────────────────

def preprocess_for_tts(name: str) -> str:
    """
    Normalise text before sending to ElevenLabs TTS.
    Only affects what gets sent to TTS — original name is still used for STT verification.

    Prepends "Hey " so the greeting + name are generated as one continuous utterance —
    avoids the seam artefact that came from splicing a separate hey.mp3.
    "&" → "and"
    """
    result = name.strip()
    result = re.sub(r"&", "and", result)
    return f"Hey {result}"


def get_voice_settings(name: str) -> dict:
    """
    Dynamic voice settings tuned to match the tone of the template part3 ("I'm the host...").

    Problem: fixed settings produced a tone that varied based on restaurant name length
    and phonetics — short names sounded clipped, long names sounded sing-songy.

    Approach:
    - stability 0.65 (was 0.4) — more consistent, less expressive variation
    - speed adjusted by approximate syllable count — longer names slow down slightly
      so they land at the same conversational pace as "I'm the host..."
    """
    syllables = len(re.findall(r"[aeiouAEIOU]+", name))

    if syllables <= 2:
        speed = 1.0    # "Nobu", "Sushi Samba" — normal pace
    elif syllables <= 5:
        speed = 0.95   # "Dishoom", "Hakkasan" — slight slowdown
    else:
        speed = 0.88   # "The Fat Duck", "Brasserie Blanc" — longer names need room

    return {
        "stability": 0.65,
        "similarity_boost": 1.0,
        "style": 0.0,
        "use_speaker_boost": True,
        "speed": speed,
    }


def generate_audio(restaurant_name: str, dest_path: Path) -> bool:
    """
    Calls ElevenLabs TTS API to generate restaurant_name in host's cloned voice.
    Pads output to exactly 2 seconds so the name is never cut abruptly.
    Saves MP3 to dest_path. Returns True on success.
    """
    api_key = os.environ["ELEVENLABS_API_KEY"]
    voice_id = os.environ["ELEVENLABS_VOICE_ID"]

    tts_text = preprocess_for_tts(restaurant_name)
    if tts_text != restaurant_name:
        log.info("  TTS text adjusted: '%s' → '%s'", restaurant_name, tts_text)

    url = ELEVENLABS_TTS_URL.format(voice_id=voice_id)
    headers = {
        "xi-api-key": api_key,
        "Content-Type": "application/json",
        "Accept": "audio/mpeg",
    }
    payload = {
        "text": tts_text,
        "model_id": ELEVENLABS_MODEL,
        "voice_settings": get_voice_settings(restaurant_name),
    }

    log.info("  Calling ElevenLabs TTS for '%s' ...", restaurant_name)
    r = requests.post(url, headers=headers, json=payload, timeout=30)
    if r.status_code == 401:
        raise RuntimeError("ElevenLabs API key invalid — check ELEVENLABS_API_KEY in .env")
    if r.status_code == 422:
        raise RuntimeError(f"ElevenLabs validation error: {r.text}")
    r.raise_for_status()

    raw_path = dest_path.with_suffix(".raw.mp3")
    raw_path.write_bytes(r.content)

    # Pad to exactly 2 seconds — ensures full phoneme is captured + natural gap before "team"
    pad_cmd = [
        "ffmpeg", "-y",
        "-i", str(raw_path),
        "-af", "apad=whole_dur=1.5",
        "-c:a", "libmp3lame",
        str(dest_path),
    ]
    pad_result = subprocess.run(pad_cmd, capture_output=True, text=True)

    if pad_result.returncode != 0:
        log.warning("  FFmpeg pad failed — using raw audio: %s", pad_result.stderr[-200:])
        dest_path.write_bytes(r.content)

    raw_path.unlink(missing_ok=True)
    log.info("  Audio saved → %s (%.1f KB)", dest_path.name, dest_path.stat().st_size / 1e3)
    return True


# ── Whisper STT verification ──────────────────────────────────────────────────

def transcribe_audio(mp3_path: Path) -> str:
    """
    Runs Whisper STT on the MP3 locally (no API cost).
    Returns the raw transcription string.
    """
    import whisper
    model = whisper.load_model(WHISPER_MODEL)
    result = model.transcribe(str(mp3_path), language="en", fp16=False)
    return result["text"].strip()


def names_match(expected: str, transcribed: str) -> bool:
    """
    Fuzzy match: normalise both strings, check containment then similarity ratio.

    Normalisation:
      - "&" → "and"  (Whisper transcribes spoken "and" not the symbol)
      - strip all remaining punctuation, lowercase

    Examples:
      expected="Dishoom"      transcribed="Dishoom."       → True  (containment)
      expected="Shrimp & Co"  transcribed="shrimp and coat" → True  (fuzzy ratio)
      expected="Le Gavroche"  transcribed="Le Gavroche team" → True  (containment)
      expected="Noma"         transcribed="Norman"          → False (ratio too low)
    """
    import difflib

    def normalise(s: str) -> str:
        s = s.replace("&", "and")
        return re.sub(r"[^a-z0-9\s]", "", s.lower()).strip()

    exp = normalise(expected)
    got = normalise(transcribed)

    # Direct containment
    if exp in got or got in exp:
        return True

    # Fuzzy similarity — catches minor mishearings like "coat" vs "co"
    ratio = difflib.SequenceMatcher(None, exp, got).ratio()
    return ratio >= 0.75


# ── Google Drive upload ───────────────────────────────────────────────────────

def get_drive_service():
    """Reuses gmail_token.json (same OAuth flow as generate_videos.py)."""
    from google.oauth2.credentials import Credentials as OAuthCredentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    DRIVE_SCOPES = [
        "https://www.googleapis.com/auth/gmail.send",
        "https://www.googleapis.com/auth/drive",
    ]
    token_file = os.environ.get("GMAIL_TOKEN_FILE", "credentials/gmail_token.json")
    creds_file = os.environ.get("GMAIL_CREDENTIALS_FILE", "credentials/oath_client.json")

    creds = None
    if Path(token_file).exists():
        creds = OAuthCredentials.from_authorized_user_file(token_file, DRIVE_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_file, DRIVE_SCOPES)
            creds = flow.run_local_server(port=0)
        Path(token_file).write_text(creds.to_json())
    return build("drive", "v3", credentials=creds)


def upload_audio_to_drive(mp3_path: Path, restaurant_name: str, folder_id: str) -> str:
    """Uploads MP3 to Drive, makes it public, returns shareable URL."""
    service = get_drive_service()
    media = MediaFileUpload(str(mp3_path), mimetype="audio/mpeg", resumable=False)
    file = service.files().create(
        body={"name": f"{restaurant_name}_name.mp3", "parents": [folder_id]},
        media_body=media,
        fields="id",
    ).execute()
    fid = file["id"]
    service.permissions().create(
        fileId=fid,
        body={"type": "anyone", "role": "reader"},
    ).execute()
    url = f"https://drive.google.com/file/d/{fid}/view?usp=sharing"
    log.info("  Uploaded audio → %s", url)
    return url


# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_sheet_client():
    key_file = os.environ["GOOGLE_SHEETS_KEY_FILE"]
    creds = SACredentials.from_service_account_file(key_file, scopes=GSHEETS_SCOPES)
    return gspread.authorize(creds)


def load_contacts(sheet_id: str):
    client = get_sheet_client()
    sheet = client.open_by_key(sheet_id)
    ws = sheet.worksheet(CONTACTS_TAB)
    all_rows = ws.get_all_values()
    if not all_rows:
        return ws, [], []
    headers = all_rows[0]
    rows = [dict(zip(headers, row)) for row in all_rows[1:]]
    return ws, headers, rows


def ensure_columns(ws: gspread.Worksheet, headers: list[str]) -> list[str]:
    """Adds drive_audio_url, audio_transcription, audio_verified columns if missing."""
    needed = ["drive_audio_url", "audio_transcription", "audio_verified"]
    new_cols = [col for col in needed if col not in headers]
    if not new_cols:
        return headers
    new_total_cols = len(headers) + len(new_cols)
    ws.resize(rows=ws.row_count, cols=new_total_cols)
    time.sleep(0.5)
    for col in new_cols:
        headers.append(col)
        ws.update_cell(1, len(headers), col)
        log.info("Added column: %s", col)
    time.sleep(1)
    return headers


def write_audio_result(
    ws: gspread.Worksheet,
    headers: list[str],
    row_index_1based: int,
    drive_audio_url: str,
    transcription: str,
    verified: bool,
):
    sheet_row = row_index_1based + 1  # +1 for header row
    ws.update_cell(sheet_row, headers.index("drive_audio_url") + 1, drive_audio_url)
    ws.update_cell(sheet_row, headers.index("audio_transcription") + 1, transcription)
    ws.update_cell(sheet_row, headers.index("audio_verified") + 1, "TRUE" if verified else "FALSE")
    status = "✓ verified" if verified else "✗ FAILED — check audio_transcription column"
    log.info("  Sheet updated (%s) | heard: '%s'", status, transcription)


# ── Core pipeline ─────────────────────────────────────────────────────────────

def run(limit: int, dry_run: bool, only_restaurant: str | None):
    sheet_id = os.environ["GOOGLE_SHEETS_ID"]
    drive_folder_id = os.environ["GOOGLE_DRIVE_FOLDER_ID"]

    log.info("Loading contacts sheet ...")
    ws, headers, rows = load_contacts(sheet_id)
    if not rows:
        log.info("No rows in sheet.")
        return
    headers = ensure_columns(ws, headers)

    # Filter: needs restaurant_name, not yet verified
    pending = []
    for i, row in enumerate(rows):
        name = row.get("restaurant_name", "").strip()
        if not name:
            continue
        if only_restaurant and name.lower() != only_restaurant.lower():
            continue
        if row.get("audio_verified", "").strip().upper() == "TRUE":
            log.debug("Row %d (%s) — already verified, skipping", i + 2, name)
            continue
        pending.append((i, row))

    log.info("%d rows to process (limit=%d)", len(pending), limit)
    if not pending:
        log.info("Nothing to do.")
        return

    processed = 0
    failed = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        for i, row in pending:
            if processed >= limit:
                log.info("Limit of %d reached.", limit)
                break

            name = row["restaurant_name"].strip()
            log.info("[%d/%d] %s", processed + 1, min(limit, len(pending)), name)

            if dry_run:
                log.info("  [dry-run] Would generate audio for '%s'", name)
                processed += 1
                continue

            # Step 1 — Generate audio via ElevenLabs TTS
            mp3_path = tmp / f"{name.replace(' ', '_')}_name.mp3"
            try:
                generate_audio(name, mp3_path)
            except Exception as exc:
                log.error("  ElevenLabs error: %s", exc)
                failed.append(name)
                continue

            # Step 2 — Whisper STT verification
            log.info("  Running Whisper STT ...")
            try:
                transcription = transcribe_audio(mp3_path)
            except Exception as exc:
                log.error("  Whisper error: %s", exc)
                failed.append(name)
                continue

            verified = names_match(name, transcription)
            log.info(
                "  Expected: '%s' | Heard: '%s' | %s",
                name, transcription, "✓ MATCH" if verified else "✗ MISMATCH",
            )

            # Step 3 — Upload MP3 to Google Drive
            try:
                drive_audio_url = upload_audio_to_drive(mp3_path, name, drive_folder_id)
            except Exception as exc:
                log.error("  Drive upload failed: %s", exc)
                failed.append(name)
                continue

            # Step 4 — Write to sheet
            write_audio_result(ws, headers, i + 1, drive_audio_url, transcription, verified)
            processed += 1

    log.info("Done. %d processed, %d failed.", processed, len(failed))
    if failed:
        log.warning("Failed restaurants: %s", failed)


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Personalised audio generator — ElevenLabs TTS + Whisper STT verification"
    )
    parser.add_argument("--limit", type=int, default=50, help="Max restaurants to process (default: 50)")
    parser.add_argument("--dry-run", action="store_true", help="Print rows without generating or uploading")
    parser.add_argument("--restaurant", metavar="NAME", help="Process a single restaurant by name")
    args = parser.parse_args()

    if args.dry_run:
        log.info("DRY RUN — no audio will be generated or uploaded")

    run(limit=args.limit, dry_run=args.dry_run, only_restaurant=args.restaurant)


if __name__ == "__main__":
    main()
