from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import os
from datetime import datetime, timedelta
import logging
import email_reader_attachment_download  # <-- Script2
import report_generator                 # <-- Script3
import tkinter as tk
import time

# Logging configuration
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


# -------------------- Countdown GUI --------------------
class CountdownGUI:
    def __init__(self, countdown_minutes=60):
        self.root = tk.Tk()
        self.root.title("Countdown to Email Fetch")
        self.root.geometry("400x200")
        self.countdown_minutes = countdown_minutes
        self.remaining = countdown_minutes * 60
        self.skip = False

        self.label = tk.Label(self.root, text="", font=("Arial", 20))
        self.label.pack(pady=20)

        self.skip_button = tk.Button(
            self.root,
            text="Skip Countdown",
            command=self.skip_countdown,
            font=("Arial", 14),
            bg="red",
            fg="white"
        )
        self.skip_button.pack(pady=10)

        self.update_label()

    def update_label(self):
        mins, secs = divmod(self.remaining, 60)
        self.label.config(text=f"Time Remaining: {mins:02}:{secs:02}")
        if self.remaining > 0 and not self.skip:
            self.remaining -= 1
            self.root.after(1000, self.update_label)
        else:
            self.root.destroy()  # ✅ Properly close GUI

    def skip_countdown(self):
        self.skip = True
        self.root.destroy()  # ✅ Close GUI instantly

    def start(self):
        self.root.mainloop()
        return self.skip


# -------------------- Shepherd Automation --------------------
class OctopusReportTester:
    def __init__(self):
        load_dotenv()
        self.username = os.getenv("OCTO_USER")
        self.password = os.getenv("OCTO_PASS")
        self.base_url = "https://octopus.eulerlogistics.com/"
        if not self.username or not self.password:
            raise ValueError("Missing credentials in environment variables")

    def login(self, page):
        logger.info("Logging in...")
        page.goto(self.base_url)
        page.wait_for_selector("input[name='username']", timeout=10000)
        page.fill("input[name='username']", self.username)
        page.fill("input[name='password']", self.password)
        page.click("button[type='submit']")
        page.wait_for_load_state("networkidle")
        logger.info("Login successful")

    def search_vehicle(self, page, registration_no):
        logger.info(f"Searching vehicle: {registration_no}")
        page.wait_for_selector("#vehicle_detail", timeout=15000)
        page.fill("#vehicle_detail", registration_no)
        page.press("#vehicle_detail", "Enter")
        page.wait_for_selector("text=Vehicle Details", timeout=15000)
        logger.info("Vehicle search completed")

    def open_shepherd_dialog(self, page):
        logger.info("Opening Shepherd Report dialog")
        download_button = page.locator(
            "div:has-text('Live Updates') button:has(svg.lucide-download)"
        ).first
        download_button.wait_for(state="visible", timeout=10000)
        download_button.click()
        shepherd_dialog = page.locator("div[role='dialog'] >> text=Shepherd Report")
        shepherd_dialog.wait_for(state="visible", timeout=10000)
        logger.info("Shepherd Report dialog opened")

    def select_start_date(self, page, start_date):
        logger.info(f"Selecting Start Date: {start_date.strftime('%Y-%m-%d')}")
        start_button = page.locator(
            'label:has-text("Start Date")'
        ).locator('xpath=following::button[@data-slot="popover-trigger"]').first
        start_button.wait_for(state="visible", timeout=5000)
        start_button.click()
        page.wait_for_timeout(500)
        self._click_date_in_calendar(page, start_date)
        page.wait_for_timeout(300)

    def select_end_date(self, page, end_date, retries=3):
        logger.info(f"Selecting End Date: {end_date.strftime('%Y-%m-%d')}")
        for attempt in range(1, retries + 1):
            try:
                end_button = page.locator("label[for='end-date']").locator("xpath=following::button[1]").first
                end_button.wait_for(state="visible", timeout=5000)
                end_button.click()
                page.wait_for_timeout(500)
                self._click_date_in_calendar(page, end_date)
                page.wait_for_timeout(300)
                logger.info("End Date selected successfully")
                return True
            except Exception as e:
                logger.warning(f"Attempt {attempt} failed: {str(e)}")
                if attempt == retries:
                    raise Exception("Could not select End Date")

    def _click_date_in_calendar(self, page, target_date):
        logger.info(f"Selecting calendar day {target_date.strftime('%Y-%m-%d')}")
        current_month = datetime.today().month
        current_year = datetime.today().year
        target_month = target_date.month
        target_year = target_date.year
        month_diff = (target_year - current_year) * 12 + (target_month - current_month)

        if month_diff > 0:
            for _ in range(month_diff):
                page.locator("button.rdp-button_next").first.click()
                page.wait_for_timeout(200)
        elif month_diff < 0:
            for _ in range(-month_diff):
                page.locator("button.rdp-button_previous").first.click()
                page.wait_for_timeout(200)

        iso_fmt = target_date.strftime("%Y-%m-%d")
        btn = page.wait_for_selector(f"td[data-day='{iso_fmt}'] button", timeout=2000)
        btn.click()
        logger.info(f"Clicked date {target_date.strftime('%Y-%m-%d')}")

    def submit_report(self, page):
        logger.info("Clicking Submit button")
        submit_button = page.locator('button[type="submit"]:has-text("Submit")').first
        submit_button.wait_for(state="visible", timeout=5000)
        submit_button.click()
        logger.info("Submit clicked successfully")

    def run_full_test(self, registration_no, start_date, end_date, headless=False):
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=headless, slow_mo=200)
            page = browser.new_page()
            page.set_extra_http_headers({'User-Agent': 'Mozilla/5.0'})
            try:
                self.login(page)
                self.search_vehicle(page, registration_no)
                self.open_shepherd_dialog(page)
                self.select_start_date(page, start_date)
                self.select_end_date(page, end_date)
                self.submit_report(page)
                logger.info("✅ Shepherd automation completed successfully")
                page.wait_for_timeout(5000)
            finally:
                browser.close()


# -------------------- Vehicle List Reader --------------------
def read_vehicle_list(filename="vehicle_list.txt"):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, filename)
    vehicles = []
    with open(file_path, "r") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                reg_no, start_date = line.split(",", 1)
                reg_no = reg_no.strip()
                start_date = start_date.replace("Start date", "").strip()
                start_date_obj = datetime.strptime(start_date, "%d%b%Y")
                vehicles.append((reg_no, start_date_obj))
            except Exception as e:
                logger.warning(f"Skipping invalid line: {line} ({e})")
    return vehicles


# -------------------- Main Flow --------------------
def main():
    tester = OctopusReportTester()
    end_date = datetime.now() - timedelta(days=1)
    vehicles = read_vehicle_list("vehicle_list.txt")

    # Phase 1: Shepherd Automation
    for reg_no, start_date in vehicles:
        logger.info(f"Processing {reg_no}: {start_date.strftime('%d-%b-%Y')} → {end_date.strftime('%d-%b-%Y')}")
        tester.run_full_test(reg_no, start_date, end_date, headless=False)

    # Countdown before Script2
    logger.info("Waiting before starting email fetch...")
    countdown = CountdownGUI(countdown_minutes=60)
    skip = countdown.start()
    if skip:
        logger.info("Countdown skipped by user.")

    # Phase 2: Email Reports Fetch
    logger.info("Starting Phase 2: Fetching reports from email...")
    email_reader_attachment_download.fetch_reports_for_all_vehicles()
    logger.info("✅ All email reports fetched successfully.")

    # Phase 3: Report Generator
    logger.info("Starting Phase 3: Generating consolidated reports...")
    # Call Script3 function with optional GUI
    report_generator.generate_all_reports(show_gui=True)
    logger.info("✅ All consolidated reports generated successfully.")


if __name__ == "__main__":
    main()
