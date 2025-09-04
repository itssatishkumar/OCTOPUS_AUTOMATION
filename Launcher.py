import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

# -----------------------------
# CONFIG
# -----------------------------
sys.path.append(r"D:\OCTOPUS_AUTOMATION")
import Octopus_login as script1
import email_reader_attachment_download as script2
import report_generator as script3

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STICKERS_DIR = os.path.join(BASE_DIR, "sticker")  # Always use "sticker"
STATE_FILE = os.path.join(BASE_DIR, "last_sticker.txt")


# -----------------------------
# Sticker Folder Manager (FIXED)
# -----------------------------
def get_sticker_folder():
    """Rotate through folders inside 'sticker' directory, robust using folder names."""
    if not os.path.exists(STICKERS_DIR):
        print(f"[ERROR] Sticker directory not found: {STICKERS_DIR}")
        return None

    all_stickers = sorted(
        [os.path.join(STICKERS_DIR, d) for d in os.listdir(STICKERS_DIR)
         if os.path.isdir(os.path.join(STICKERS_DIR, d))]
    )

    if not all_stickers:
        print("[ERROR] No subfolders found inside 'sticker'")
        return None

    basenames = [os.path.basename(p) for p in all_stickers]

    last_value = None
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                last_value = f.read().strip()
        except Exception:
            last_value = None

    # Default to first folder
    current_index = 0

    if last_value and last_value in basenames:
        idx = basenames.index(last_value)
        current_index = (idx + 1) % len(all_stickers)

    chosen_folder = all_stickers[current_index]

    # Save the folder name instead of index
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        f.write(os.path.basename(chosen_folder))

    print(f"[INFO] Using sticker folder: {chosen_folder}")
    return chosen_folder


def load_frames(folder, max_size=(200, 200)):
    """Load sticker frames and resize proportionally."""
    frames = []
    if folder and os.path.exists(folder):
        for file in sorted(os.listdir(folder)):
            if file.lower().endswith((".png", ".webp")):
                image_path = os.path.join(folder, file)
                image = Image.open(image_path)
                image.thumbnail(max_size, Image.Resampling.LANCZOS)
                frames.append(ImageTk.PhotoImage(image))
        print(f"[INFO] Loaded {len(frames)} frames from {folder}")
    else:
        print(f"[ERROR] Folder not found: {folder}")
    return frames


# -----------------------------
# Professional Launcher GUI
# -----------------------------
class LauncherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Octopus Automation Launcher")
        self.root.geometry("800x500")
        self.root.configure(bg="#f5f5f5")
        self.root.resizable(False, False)

        # Main layout
        main_frame = tk.Frame(root, bg="#f5f5f5")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # =========================
        # LEFT PANEL (Buttons/Text)
        # =========================
        left_frame = tk.Frame(main_frame, bg="#f5f5f5")
        left_frame.pack(side="left", fill="y", padx=(0, 20))

        title = tk.Label(left_frame, text="Octopus Automation Launcher",
                         font=("Segoe UI", 20, "bold"), bg="#f5f5f5", fg="#333")
        title.pack(pady=10)

        subtitle = tk.Label(left_frame, text="Select a script to run:",
                            font=("Segoe UI", 12), bg="#f5f5f5", fg="#555")
        subtitle.pack(pady=(0, 20))

        self.btn1 = tk.Button(left_frame, text="Run Octopus Login",
                              font=("Segoe UI", 12, "bold"), bg="#4CAF50", fg="white",
                              width=25, command=lambda: self.run_script(1))
        self.btn1.pack(pady=10)

        self.btn2 = tk.Button(left_frame, text="Run Email Fetch",
                              font=("Segoe UI", 12, "bold"), bg="#2196F3", fg="white",
                              width=25, command=lambda: self.run_script(2))
        self.btn2.pack(pady=10)

        self.btn3 = tk.Button(left_frame, text="Generate Reports",
                              font=("Segoe UI", 12, "bold"), bg="#FF9800", fg="white",
                              width=25, command=lambda: self.run_script(3))
        self.btn3.pack(pady=10)

        self.status_label = tk.Label(left_frame, text="Ready", font=("Segoe UI", 11),
                                     bg="#f5f5f5", fg="#555")
        self.status_label.pack(pady=20)

        # =========================
        # RIGHT PANEL (Sticker)
        # =========================
        right_frame = tk.Frame(main_frame, bg="#f5f5f5")
        right_frame.pack(side="right", fill="both", expand=True)

        sticker_folder = get_sticker_folder()
        self.frames = load_frames(sticker_folder)

        if self.frames:
            self.sticker_label = tk.Label(right_frame, bg="#f5f5f5")
            self.sticker_label.pack(expand=True)
            self.animate_sticker()
        else:
            tk.Label(right_frame, text="No sticker found!", font=("Segoe UI", 12, "italic"),
                     bg="#f5f5f5", fg="red").pack(expand=True)

    # -----------------------------
    # Sticker Animation
    # -----------------------------
    def animate_sticker(self, idx=0):
        if self.frames:
            frame = self.frames[idx]
            self.sticker_label.config(image=frame)
            next_idx = (idx + 1) % len(self.frames)
            self.root.after(100, self.animate_sticker, next_idx)

    # -----------------------------
    # Script Runner
    # -----------------------------
    def run_script(self, script_number):
        self.disable_buttons()
        if script_number == 1:
            self.status_label.config(text="Running Script 1 → Script 2 → Script 3 …")
            threading.Thread(target=self.run_script1_flow, daemon=True).start()
        elif script_number == 2:
            self.status_label.config(text="Running Script 2 → Script 3 …")
            threading.Thread(target=self.run_script2_flow, daemon=True).start()
        elif script_number == 3:
            self.status_label.config(text="Running Script 3 …")
            threading.Thread(target=self.run_script3_only, daemon=True).start()

    def disable_buttons(self):
        self.btn1.config(state="disabled")
        self.btn2.config(state="disabled")
        self.btn3.config(state="disabled")

    def enable_buttons(self):
        self.btn1.config(state="normal")
        self.btn2.config(state="normal")
        self.btn3.config(state="normal")

    # -----------------------------
    # Script Flows
    # -----------------------------
    def run_script1_flow(self):
        try:
            script1.main()
            script2.fetch_reports_for_all_vehicles()
            script3.generate_all_reports()
            messagebox.showinfo("Completed", "Script 1 → Script 2 → Script 3 completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.enable_buttons()
            self.status_label.config(text="Ready")

    def run_script2_flow(self):
        try:
            script2.fetch_reports_for_all_vehicles()
            script3.generate_all_reports()
            messagebox.showinfo("Completed", "Script 2 → Script 3 completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.enable_buttons()
            self.status_label.config(text="Ready")

    def run_script3_only(self):
        try:
            script3.generate_all_reports()
            messagebox.showinfo("Completed", "Script 3 completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.enable_buttons()
            self.status_label.config(text="Ready")


# -----------------------------
# MAIN
# -----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = LauncherGUI(root)
    root.mainloop()
