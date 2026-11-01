import os
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

WORKBOOK_FILE = os.path.join(BASE_DIR, "Deviation_Tracker.xlsx")
INCOMING_FOLDER = os.path.join(BASE_DIR, "incoming")
PROCESSED_FOLDER = os.path.join(BASE_DIR, "processed")

SEND_MODE = os.getenv("SEND_MODE") or "DRYRUN"
TEST_RECIPIENT = os.getenv("TEST_RECIPIENT")
TRACKER_LOCATION = os.getenv("TRACKER_LOCATION") or WORKBOOK_FILE

GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

PICKUP_MODE = os.getenv("PICKUP_MODE") or "FOLDER"
LOOKER_SENDER = os.getenv("LOOKER_SENDER") or ""
LOOKER_SUBJECT = os.getenv("LOOKER_SUBJECT") or ""
IMAP_HOST = os.getenv("IMAP_HOST") or "imap.gmail.com"