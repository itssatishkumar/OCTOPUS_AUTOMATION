#!/usr/bin/env python3
"""
Script 4: Google Drive Temp Tracker Upload using OAuth2 (Personal Drive)
Standalone version with automated token reuse
"""

import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# =========================
# CONFIGURATION
# =========================
DOWNLOAD_FOLDER = r"D:\OCTOPUS_AUTOMATION\download"  # Root vehicle folders
OAUTH_JSON_FILE = r"D:\OCTOPUS_AUTOMATION\ntc_tracker_oauth.json"
TOKEN_FILE = r"D:\OCTOPUS_AUTOMATION\token.json"    # OAuth token will be saved here
MAX_CONCURRENT_UPLOADS = 2  # Simultaneous uploads
AUTO_CLOSE_SECS = 2  # GUI auto-close after uploads
ROOT_DRIVE_FOLDER_NAME = "NTC TRACKER"  # Root folder in Google Drive

SCOPES = ['https://www.googleapis.com/auth/drive.file']  # Limited access to personal drive

# =========================
# GUI
# =========================
class UploadProgressGUI:
    def __init__(self, vehicle_names):
        self.root = tk.Tk()
        self.root.title("NTC Tracker Upload")
        self.root.geometry("820x500")
        tk.Label(self.root, text="Google Drive Upload Progress", font=("Arial", 16, "bold")).pack(pady=(8,0))

        container = tk.Frame(self.root)
        container.pack(fill="both", expand=True, padx=8, pady=(6,2))

        self.canvas = tk.Canvas(container, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0,0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.vehicle_widgets = {}
        for name in vehicle_names:
            row = tk.Frame(self.scrollable_frame, pady=6)
            row.pack(fill="x", padx=6)
            lbl_name = tk.Label(row, text=name, font=("Arial", 10, "bold"), width=22, anchor="w")
            lbl_name.pack(side="left", padx=(2,6))
            pb = ttk.Progressbar(row, orient="horizontal", length=300, mode="determinate")
            pb.pack(side="left", padx=(0,8))
            lbl_status = tk.Label(row, text="Waiting…", font=("Arial",9), anchor="w", wraplength=350, justify="left")
            lbl_status.pack(side="left", fill="x", expand=True)
            self.vehicle_widgets[name] = {"progress": pb, "status": lbl_status, "name_label": lbl_name}

        footer = tk.Frame(self.root)
        footer.pack(fill="x", pady=(4,8))
        self.global_status = tk.Label(footer, text="Ready", font=("Arial",11))
        self.global_status.pack(side="left", padx=8)
        self.close_btn = tk.Button(footer, text="Close Now", command=self._close_now, state="disabled")
        self.close_btn.pack(side="right", padx=8)
        self._closed = False

    def start(self):
        self.root.mainloop()

    def _close_now(self):
        self._closed = True
        try:
            self.root.destroy()
        except:
            pass

    def update_progress(self, vehicle_name, percent, status):
        def _apply():
            if vehicle_name in self.vehicle_widgets:
                w = self.vehicle_widgets[vehicle_name]
                w["progress"]["value"] = max(0, min(100, percent))
                w["status"].config(text=status)
                self.global_status.config(text=f"Processing: {vehicle_name} — {status}")
        try:
            self.root.after(0, _apply)
        except:
            pass

    def mark_vehicle_done(self, vehicle_name, msg="Done ✅"):
        def _apply():
            if vehicle_name in self.vehicle_widgets:
                w = self.vehicle_widgets[vehicle_name]
                w["progress"]["value"] = 100
                w["status"].config(text=msg)
                self.global_status.config(text=f"Completed: {vehicle_name}")
        try:
            self.root.after(0, _apply)
        except:
            pass

    def mark_done(self, msg="✅ All uploads completed"):
        def _apply():
            self.global_status.config(text=msg)
            self.close_btn.config(state="normal")
            if not self._closed and AUTO_CLOSE_SECS:
                self.root.after(int(AUTO_CLOSE_SECS*1000), self._close_now)
        try:
            self.root.after(0, _apply)
        except:
            pass

# =========================
# Google Drive Helpers
# =========================
def get_drive_service():
    creds = None
    # Load token if exists
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # If no valid credentials, run OAuth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(OAUTH_JSON_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save credentials for next time
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    service = build('drive', 'v3', credentials=creds)
    return service

def get_or_create_folder(service, parent_id, folder_name):
    query = f"mimeType='application/vnd.google-apps.folder' and trashed=false and name='{folder_name}' and '{parent_id}' in parents"
    res = service.files().list(q=query, fields="files(id, name)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    metadata = {"name": folder_name, "mimeType": "application/vnd.google-apps.folder", "parents":[parent_id]}
    folder = service.files().create(body=metadata, fields="id").execute()
    return folder["id"]

def delete_existing_docs(service, folder_id):
    query = f"'{folder_id}' in parents and trashed=false and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'"
    res = service.files().list(q=query, fields="files(id, name)").execute()
    for f in res.get("files", []):
        service.files().delete(fileId=f["id"]).execute()

def upload_docx(service, folder_id, local_file_path):
    media = MediaFileUpload(local_file_path, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    fname = os.path.basename(local_file_path)
    service.files().create(body={"name": fname, "parents":[folder_id]}, media_body=media, fields="id").execute()

# =========================
# Per-Vehicle Upload
# =========================
def upload_vehicle_report(vehicle_folder, service, root_folder_id, progress_cb=None):
    vehicle_name = os.path.basename(vehicle_folder)
    report_files = [f for f in os.listdir(vehicle_folder) if f.lower().endswith(".docx") and f.startswith("temp_report_")]

    vehicle_folder_id = get_or_create_folder(service, root_folder_id, vehicle_name)

    if not report_files:
        if progress_cb:
            progress_cb(vehicle_name, 0, "Processed, but no report ❌")
        return

    report_file = os.path.join(vehicle_folder, report_files[0])
    if progress_cb:
        progress_cb(vehicle_name, 5, "Preparing upload…")

    delete_existing_docs(service, vehicle_folder_id)
    if progress_cb:
        progress_cb(vehicle_name, 50, "Old files deleted…")

    upload_docx(service, vehicle_folder_id, report_file)
    if progress_cb:
        progress_cb(vehicle_name, 100, f"Uploaded {os.path.basename(report_file)} ✅")

# =========================
# Main Batch Upload
# =========================
def main():
    if not os.path.isdir(DOWNLOAD_FOLDER):
        print(f"❌ Download folder not found: {DOWNLOAD_FOLDER}")
        return

    vehicle_folders = [os.path.join(DOWNLOAD_FOLDER, d) for d in os.listdir(DOWNLOAD_FOLDER)
                       if os.path.isdir(os.path.join(DOWNLOAD_FOLDER, d))]
    if not vehicle_folders:
        print("❌ No vehicle folders found")
        return

    vehicle_names = [os.path.basename(v) for v in vehicle_folders]
    gui = UploadProgressGUI(vehicle_names)

    service = get_drive_service()
    root_folder_id = get_or_create_folder(service, "root", ROOT_DRIVE_FOLDER_NAME)

    def progress_wrapper(vname, pct, status):
        gui.update_progress(vname, pct, status)

    def task(folder):
        upload_vehicle_report(folder, service, root_folder_id, progress_cb=progress_wrapper)
        gui.mark_vehicle_done(os.path.basename(folder))

    def run_executor():
        with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_UPLOADS) as executor:
            futures = [executor.submit(task, f) for f in vehicle_folders]
            for f in as_completed(futures):
                try:
                    f.result()
                except Exception as e:
                    print("❌ Error:", e)
        gui.mark_done("✅ All uploads completed")

    t = threading.Thread(target=run_executor, daemon=True)
    t.start()
    gui.start()

if __name__ == "__main__":
    main()
