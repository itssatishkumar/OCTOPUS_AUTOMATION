# email_reader_attachment_download.py
import imaplib
import email
import os
import logging
import datetime
import requests
import threading
import time
import tkinter as tk
from email.header import decode_header
from dotenv import load_dotenv
from bs4 import BeautifulSoup  # pip install beautifulsoup4
import pandas as pd            # pip install pandas openpyxl
import re

# ----------------- Config -----------------
DEFAULT_WAIT_MINUTES = 60  # Default countdown (in minutes)
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993

# Logging setup
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

ROOT_DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "download")


# ----------------- Helpers -----------------
def decode_mime_words(s):
    if not s:
        return ""
    decoded = decode_header(s)
    subject = ""
    for part, enc in decoded:
        if isinstance(part, bytes):
            subject += part.decode(enc or "utf-8", errors="ignore")
        else:
            subject += part
    return subject


def connect_to_mailbox():
    """Connect to IMAP mailbox."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")
        return mail
    except Exception as e:
        logger.error(f"Failed to connect to mailbox: {e}")
        raise


def search_latest_vehicle_email(mail, vehicle_id):
    """Find latest email with Internal Reports for a vehicle."""
    since_date = (datetime.datetime.now() - datetime.timedelta(days=2)).strftime("%d-%b-%Y")
    search_query = f'(SINCE "{since_date}" SUBJECT "Internal Reports {vehicle_id}")'
    status, data = mail.search(None, search_query)

    if status != "OK" or not data[0]:
        logger.warning(f"No emails found for {vehicle_id}. Raw search data: {data}")
        return None

    email_ids = data[0].split()
    logger.info(f"Found {len(email_ids)} emails for {vehicle_id}. Using latest one.")
    return email_ids[-1]


def clean_or_create_folder(vehicle_id):
    """Ensure folder exists. Do NOT delete old files."""
    vehicle_folder = os.path.join(ROOT_DOWNLOAD_DIR, vehicle_id)
    os.makedirs(vehicle_folder, exist_ok=True)
    return vehicle_folder


def parse_email_date(date_str):
    """
    Parse a date like '1/9/2025' (DD/MM/YYYY) to a date object.
    If the cell contains extra text, extract the first DD/MM/YYYY pattern.
    Returns datetime.date or None.
    """
    # Try direct parse (handles many formats)
    dt = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        return dt.date()

    # Fallback: regex extract first dd/mm/yyyy
    m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', str(date_str))
    if m:
        dt = pd.to_datetime(m.group(1), dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    return None


def get_existing_first_created_dates(folder):
    """
    For each .xlsx in folder, read the first valid (top-most) createdAt value.
    Use only that date (date part only) per file.
    Returns a set of datetime.date.
    """
    existing_dates = set()
    for file in os.listdir(folder):
        if file.lower().endswith(".xlsx"):
            path = os.path.join(folder, file)
            try:
                df = pd.read_excel(path, engine="openpyxl")
                if "createdAt" in df.columns:
                    ser = df["createdAt"].dropna()
                    if not ser.empty:
                        # First non-null value only (as requested)
                        first_val = ser.iloc[0]
                        dt = pd.to_datetime(first_val, errors="coerce")
                        if pd.notna(dt):
                            existing_dates.add(dt.date())
                        else:
                            logger.debug(f"First createdAt not a valid datetime in {file}: {first_val}")
                else:
                    logger.debug(f"'createdAt' column not found in {file}")
            except Exception as e:
                logger.warning(f"Could not read {path}: {e}")
    return existing_dates


def extract_all_links(msg):
    """
    Extract one report link per date row.
    Priority: CAN > CSV.
    """
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() in ["text/plain", "text/html"]:
                try:
                    body += part.get_payload(decode=True).decode(errors="ignore")
                except:
                    pass
    else:
        try:
            body = msg.get_payload(decode=True).decode(errors="ignore")
        except:
            body = ""

    soup = BeautifulSoup(body, "html.parser")
    links = []

    # Table rows with Date | CSV Report | CAN Report
    for row in soup.find_all("tr"):
        cols = row.find_all("td")
        if len(cols) < 3:
            continue

        date_range = cols[0].get_text(strip=True)

        # CAN (priority)
        can_cell = cols[2]
        can_link = can_cell.find("a")
        if can_link and can_link.has_attr("href"):
            links.append((can_link["href"], date_range, "can"))
            continue  # skip CSV if CAN exists

        # CSV (fallback)
        csv_cell = cols[1]
        csv_link = csv_cell.find("a")
        if csv_link and csv_link.has_attr("href"):
            links.append((csv_link["href"], date_range, "csv"))

    return links


def download_file(url, folder, date_range=None):
    """Download file from given URL."""
    base_name = url.split("/")[-1].split("?")[0]
    if date_range:
        base_name = f"{date_range.replace('/', '-')}_{base_name}"
    local_filename = os.path.join(folder, base_name)

    logger.info(f"Downloading: {url}")
    try:
        resp = requests.get(url, stream=True, timeout=30)
        resp.raise_for_status()
        with open(local_filename, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        logger.info(f"Saved file: {local_filename}")
        return local_filename
    except Exception as e:
        logger.error(f"Failed to download {url}: {e}")
        return None


def process_vehicle(mail, vehicle_id):
    """Process a single vehicle email and download reports."""
    logger.info(f"Processing vehicle: {vehicle_id}")
    email_id = search_latest_vehicle_email(mail, vehicle_id)
    if not email_id:
        return

    status, msg_data = mail.fetch(email_id, "(RFC822)")
    if status != "OK":
        logger.error(f"Failed to fetch email for {vehicle_id}")
        return

    raw_msg = msg_data[0][1]
    msg = email.message_from_bytes(raw_msg)

    vehicle_folder = clean_or_create_folder(vehicle_id)
    existing_first_dates = get_existing_first_created_dates(vehicle_folder)

    links = extract_all_links(msg)
    if not links:
        logger.warning(f"No report links found for {vehicle_id}")
        return

    for idx, (link, date_range, report_type) in enumerate(links, 1):
        link_date = parse_email_date(date_range)
        if not link_date:
            logger.warning(f"Invalid date in email row for {vehicle_id}: {date_range}")
            continue

        if link_date in existing_first_dates:
            logger.info(f"Skipping {vehicle_id} report for {link_date} (already present by first createdAt in some xlsx)")
            continue

        logger.info(
            f"Downloading {report_type.upper()} report {idx}/{len(links)} "
            f"for {vehicle_id} (Date: {date_range})"
        )
        download_file(link, vehicle_folder, date_range)


def fetch_reports_for_all_vehicles(vehicle_file="vehicle_list.txt"):
    """Fetch reports for all vehicles listed in the text file."""
    logger.info("Starting Phase 2: Fetching reports from email...")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, vehicle_file)

    if not os.path.exists(file_path):
        logger.error(f"{vehicle_file} not found.")
        return

    with open(file_path, "r") as f:
        vehicle_ids = [line.split(",")[0].strip() for line in f if line.strip()]

    if not vehicle_ids:
        logger.warning(f"{vehicle_file} is empty. Nothing to process.")
        return

    mail = connect_to_mailbox()
    try:
        for vehicle_id in vehicle_ids:
            try:
                process_vehicle(mail, vehicle_id)
            except Exception as e:
                logger.error(f"Error processing {vehicle_id}: {e}")
                continue
    finally:
        mail.logout()

    logger.info("✅ All email reports fetched successfully.")


# ----------------- Countdown GUI -----------------
def start_countdown(wait_minutes=DEFAULT_WAIT_MINUTES):
    """Show a countdown GUI before running Script 2."""
    def countdown():
        nonlocal remaining
        while remaining > 0 and not skip_event.is_set():
            mins, secs = divmod(remaining, 60)
            timer_label.config(text=f"Starting in {mins:02d}:{secs:02d}")
            time.sleep(1)
            remaining -= 1
        root.destroy()
        run_script()

    def skip():
        skip_event.set()
        root.destroy()
        run_script()

    skip_event = threading.Event()
    remaining = wait_minutes * 60

    root = tk.Tk()
    root.title("Countdown Before Script2")
    root.geometry("320x150")
    root.resizable(False, False)

    timer_label = tk.Label(root, text="", font=("Arial", 18))
    timer_label.pack(pady=20)

    skip_button = tk.Button(root, text="Skip Countdown", command=skip, font=("Arial", 12))
    skip_button.pack(pady=10)

    threading.Thread(target=countdown, daemon=True).start()
    root.mainloop()


def run_script():
    fetch_reports_for_all_vehicles()


if __name__ == "__main__":
    start_countdown()
