import email
import imaplib
import os
from email.header import decode_header
from email.utils import parsedate_to_datetime

import pandas as pd

import settings

STATE_FILE = os.path.join(settings.BASE_DIR, "mail_state.txt")
LOOKBACK_DAYS = 14


def decode_text(value):
    if not value:
        return ""
    parts = decode_header(value)
    return "".join(
        chunk.decode(charset or "utf-8", errors="replace")
        if isinstance(chunk, bytes) else chunk
        for chunk, charset in parts
    )


def load_state():
    if not os.path.exists(STATE_FILE):
        return set()
    with open(STATE_FILE, encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def remember(message_id):
    with open(STATE_FILE, "a", encoding="utf-8") as f:
        f.write(message_id + "\n")


def received_time(msg):
    try:
        return parsedate_to_datetime(msg["Date"])
    except Exception:
        return pd.Timestamp.now()


def matching_messages(imap):
    since = (pd.Timestamp.today() - pd.Timedelta(days=LOOKBACK_DAYS)).strftime("%d-%b-%Y")
    criteria = ["SINCE", since]
    if settings.LOOKER_SENDER:
        criteria += ["FROM", f'"{settings.LOOKER_SENDER}"']
    if settings.LOOKER_SUBJECT:
        criteria += ["SUBJECT", f'"{settings.LOOKER_SUBJECT}"']

    imap.select("INBOX", readonly=True)
    status, data = imap.search(None, *criteria)
    for number in data[0].split():
        status, msg_data = imap.fetch(number, "(BODY.PEEK[])")
        yield email.message_from_bytes(msg_data[0][1])


def fetch_new_pdfs(preview=False):
    if not settings.GMAIL_ADDRESS or not settings.GMAIL_APP_PASSWORD:
        raise ValueError("GMAIL_ADDRESS or GMAIL_APP_PASSWORD missing in .env")

    os.makedirs(settings.INCOMING_FOLDER, exist_ok=True)
    seen = load_state()
    saved = []

    imap = imaplib.IMAP4_SSL(settings.IMAP_HOST)
    imap.login(settings.GMAIL_ADDRESS, settings.GMAIL_APP_PASSWORD)
    try:
        for msg in matching_messages(imap):
            received = received_time(msg)
            subject = decode_text(msg["Subject"])
            message_id = (msg.get("Message-ID") or f"{received}|{subject}").strip()
            if message_id in seen:
                continue

            found_pdf = False
            for part in msg.walk():
                filename = decode_text(part.get_filename())
                if not filename.lower().endswith(".pdf"):
                    continue
                found_pdf = True
                target_name = f"{received:%Y%m%d_%H%M%S}_{os.path.basename(filename)}"
                if preview:
                    print(f"Would save '{target_name}' from email '{subject}' "
                          f"(received {received:%d %b %Y %H:%M})")
                else:
                    path = os.path.join(settings.INCOMING_FOLDER, target_name)
                    with open(path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    saved.append(path)

            if found_pdf and not preview:
                remember(message_id)
    finally:
        imap.logout()

    return saved


if __name__ == "__main__":
    fetch_new_pdfs(preview=True)
    print("Preview only: nothing was saved.")