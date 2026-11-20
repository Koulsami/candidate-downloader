"""
Candidate Attachment Downloader — Mac version
---------------------------------------------
Downloads attachments (e.g. LinkedIn candidate profiles) from a chosen
mail folder and saves them into a destination folder of your choice,
renamed as:  <Candidate Name> - <original filename>.<ext>

How it works on Mac:
  Instead of automating the Outlook app (unreliable on macOS), this app
  connects DIRECTLY to your Gmail account over IMAP. The folders you see
  in Outlook are Gmail labels, so the exact same folder will appear in
  the folder list here.

Requirements:
  - macOS with Python 3.9+  (python.org installer includes tkinter)
  - A Google "App Password" (see setup instructions)
  - No third-party packages needed — 100% standard library.

Run:  python3 candidate_attachment_downloader_mac.py
"""

import os
import re
import json
import email
import imaplib
import threading
from email.header import decode_header, make_header
from email.utils import parseaddr
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

IMAP_HOST = "imap.gmail.com"

ALLOWED_EXTS = {".pdf", ".doc", ".docx", ".rtf", ".odt"}

SUBJECT_PATTERNS = [
    re.compile(r"new application[:\s]+(.+?)(?:\s*[-–|].*)?$", re.I),
    re.compile(r"application from\s+(.+?)(?:\s*[-–|].*)?$", re.I),
    re.compile(r"(.+?)\s+applied (?:for|to)\b", re.I),
    re.compile(r"(?:your applicant|applicant)[:\s]+(.+?)$", re.I),
    re.compile(r"(.+?)\s*[-–]\s*(?:resume|cv|profile)\b", re.I),
]

LOG_FILENAME = "_downloaded_log.json"
INVALID_FS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def sanitize(name: str) -> str:
    name = INVALID_FS_CHARS.sub("", name).strip().strip(".")
    return re.sub(r"\s+", " ", name) or "Unknown"


def decode_hdr(value) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)


def guess_candidate_name(subject: str, sender_name: str, sender_email: str) -> str:
    subject = (subject or "").strip()
    for pat in SUBJECT_PATTERNS:
        m = pat.search(subject)
        if m:
            cand = m.group(1).strip()
            if 1 < len(cand) <= 60:
                return sanitize(cand)
    sname = (sender_name or "").strip()
    if sname and "linkedin" not in sname.lower():
        return sanitize(sname)
    if sender_email and "@" in sender_email:
        return sanitize(sender_email.split("@")[0])
    return "Unknown"


def imap_utf7_decode(s: str) -> str:
    """Decode IMAP modified UTF-7 folder names for display."""
    try:
        def repl(m):
            b = m.group(1)
            if not b:
                return "&"
            return (b.replace(",", "/") + "===").encode().decode("utf-7").lstrip("+")
        return re.sub(r"&([^-]*)-", lambda m: repl(m), s)
    except Exception:
        return s


# ---------------------------------------------------------------------------
# GUI application
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Candidate Attachment Downloader (Mac)")
        self.geometry("940x680")
        self.minsize(780, 560)

        self.imap = None
        self.folder_names = {}     # display name -> raw IMAP name
        self.scanned_rows = []     # dicts: uid, candidate, attachment, subject, received
        self.current_folder_raw = None

        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        # --- login row ---
        login = ttk.LabelFrame(top, text="Gmail connection", padding=8)
        login.pack(fill="x")
        ttk.Label(login, text="Gmail address:").grid(row=0, column=0, sticky="w")
        self.email_var = tk.StringVar()
        ttk.Entry(login, textvariable=self.email_var, width=32).grid(
            row=0, column=1, sticky="w", padx=(4, 16))
        ttk.Label(login, text="App password:").grid(row=0, column=2, sticky="w")
        self.pw_var = tk.StringVar()
        ttk.Entry(login, textvariable=self.pw_var, show="•", width=22).grid(
            row=0, column=3, sticky="w", padx=4)
        ttk.Button(login, text="Connect", command=self.connect).grid(
            row=0, column=4, padx=8)

        # --- folder + destination ---
        ttk.Label(top, text="Mail folder (same folders you see in Outlook):").pack(
            anchor="w", pady=(10, 0))
        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(
            top, textvariable=self.folder_var, state="readonly")
        self.folder_combo.pack(fill="x", pady=(2, 8))

        ttk.Label(top, text="Save attachments to (destination folder):").pack(anchor="w")
        dest_row = ttk.Frame(top)
        dest_row.pack(fill="x", pady=(2, 8))
        self.dest_var = tk.StringVar()
        ttk.Entry(dest_row, textvariable=self.dest_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(dest_row, text="Browse…", command=self._pick_dest).pack(
            side="left", padx=6)

        opts = ttk.Frame(top)
        opts.pack(fill="x")
        self.skip_done_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts, text="Skip emails already downloaded previously",
            variable=self.skip_done_var).pack(side="left")
        ttk.Label(opts, text="   Scan emails from the last").pack(side="left")
        self.days_var = tk.StringVar(value="90")
        ttk.Entry(opts, textvariable=self.days_var, width=5).pack(side="left", padx=3)
        ttk.Label(opts, text="days").pack(side="left")

        btns = ttk.Frame(top)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="1. Scan folder", command=self.scan).pack(side="left")
        ttk.Button(btns, text="2. Download attachments",
                   command=self.download).pack(side="left", padx=8)

        # --- preview table ---
        mid = ttk.Frame(self, padding=(10, 4, 10, 0))
        mid.pack(fill="both", expand=True)
        ttk.Label(mid, text="Preview  (double-click a Candidate Name to edit it "
                            "before downloading):").pack(anchor="w")
        cols = ("candidate", "attachment", "subject", "received")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings")
        for cid, text, width in [
            ("candidate", "Candidate Name (editable)", 220),
            ("attachment", "Attachment", 200),
            ("subject", "Email Subject", 280),
            ("received", "Received", 120),
        ]:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=width, anchor="w")
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        self.tree.bind("<Double-1>", self._edit_cell)

        self.status = tk.StringVar(value="Not connected. Enter your Gmail address "
                                         "and App Password, then click Connect.")
        ttk.Label(self, textvariable=self.status, padding=8,
                  relief="sunken", anchor="w").pack(fill="x", side="bottom")

    # ---------------- Connect ----------------
    def connect(self):
        addr = self.email_var.get().strip()
        pw = self.pw_var.get().strip().replace(" ", "")
        if not addr or not pw:
            messagebox.showwarning("Missing details",
                                   "Enter your Gmail address and App Password.")
            return
        self.status.set("Connecting to Gmail…")
        self.update_idletasks()
        try:
            self.imap = imaplib.IMAP4_SSL(IMAP_HOST)
            self.imap.login(addr, pw)
            typ, data = self.imap.list()
            names = {}
            for line in data:
                if not line:
                    continue
                s = line.decode(errors="replace")
                m = re.search(r'"([^"]+)"$|\s(\S+)$', s)
                raw = (m.group(1) or m.group(2)) if m else None
                if not raw:
                    continue
                if raw.startswith("[Gmail]") and raw != "[Gmail]/All Mail":
                    # keep All Mail, hide Spam/Trash/etc. noise? keep everything visible:
                    pass
                display = imap_utf7_decode(raw)
                names[display] = raw
            self.folder_names = names
            self.folder_combo["values"] = sorted(names.keys())
            self.status.set(f"Connected — {len(names)} folders/labels found. "
                            "Pick the candidate folder and click Scan.")
        except imaplib.IMAP4.error as e:
            messagebox.showerror(
                "Login failed",
                "Gmail rejected the login.\n\n"
                "• Make sure you are using an App Password (16 characters), "
                "not your normal Gmail password.\n"
                "• App Passwords require 2-Step Verification to be ON.\n\n"
                f"Details: {e}")
            self.status.set("Login failed.")
        except Exception as e:
            messagebox.showerror("Connection failed", str(e))
            self.status.set("Connection failed.")

    def _pick_dest(self):
        d = filedialog.askdirectory(title="Choose destination folder")
        if d:
            self.dest_var.set(d)

    # ---------------- Scan (threaded) ----------------
    def scan(self):
        if self.imap is None:
            messagebox.showwarning("Not connected", "Connect to Gmail first.")
            return
        disp = self.folder_var.get()
        if not disp:
            messagebox.showwarning("Pick a folder", "Choose the mail folder first.")
            return
        threading.Thread(target=self._scan_worker,
                         args=(self.folder_names[disp],), daemon=True).start()

    def _scan_worker(self, raw_folder):
        self.current_folder_raw = raw_folder
        try:
            days = max(1, int(self.days_var.get()))
        except ValueError:
            days = 90
        since = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")

        done_ids = set()
        dest = self.dest_var.get()
        if self.skip_done_var.get() and dest:
            done_ids = self._load_log(dest)

        def ui(fn, *a):
            self.after(0, fn, *a)

        ui(self.status.set, "Scanning folder…")
        ui(self.tree.delete, *self.tree.get_children())
        self.scanned_rows = []

        try:
            quoted = f'"{raw_folder}"' if " " in raw_folder else raw_folder
            typ, _ = self.imap.select(quoted, readonly=True)
            if typ != "OK":
                ui(self.status.set, f"Could not open folder {raw_folder}.")
                return
            typ, data = self.imap.uid("SEARCH", None, f"(SINCE {since})")
            uids = data[0].split() if data and data[0] else []
            uids = [u.decode() for u in uids]
            uids = [u for u in uids if f"{raw_folder}:{u}" not in done_ids]

            found = 0
            for i, uid in enumerate(reversed(uids)):  # newest first
                ui(self.status.set,
                   f"Scanning… {i + 1}/{len(uids)} emails checked, "
                   f"{found} attachment(s) found")
                typ, msg_data = self.imap.uid("FETCH", uid, "(BODY.PEEK[])")
                if typ != "OK" or not msg_data or msg_data[0] is None:
                    continue
                raw_bytes = msg_data[0][1]
                msg = email.message_from_bytes(raw_bytes)

                sender_name, sender_email = parseaddr(decode_hdr(msg.get("From")))
                subject = decode_hdr(msg.get("Subject"))
                try:
                    received = email.utils.parsedate_to_datetime(
                        msg.get("Date")).strftime("%Y-%m-%d %H:%M")
                except Exception:
                    received = ""
                cand = guess_candidate_name(subject, sender_name, sender_email)

                att_idx = 0
                for part in msg.walk():
                    fname = part.get_filename()
                    if not fname:
                        continue
                    fname = decode_hdr(fname)
                    ext = os.path.splitext(fname)[1].lower()
                    att_idx += 1
                    if ext not in ALLOWED_EXTS:
                        continue
                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue
                    row = {
                        "uid": uid, "folder": raw_folder,
                        "candidate": cand, "attachment": fname,
                        "subject": subject, "received": received,
                        "data": payload,
                    }
                    self.scanned_rows.append(row)
                    found += 1
                    ui(self._insert_row, cand, fname, subject, received)

            ui(self.status.set,
               f"Scan complete — {found} attachment(s) found in the last "
               f"{days} days" + (" (already-downloaded emails skipped)." if done_ids else "."))
        except Exception as e:
            ui(self.status.set, f"Scan error: {e}")

    def _insert_row(self, *vals):
        self.tree.insert("", "end", values=vals)

    # ---------------- Inline editing ----------------
    def _edit_cell(self, event):
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id or col != "#1":
            return
        x, y, w, h = self.tree.bbox(row_id, col)
        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, self.tree.set(row_id, "candidate"))
        entry.focus()

        def save(_=None):
            new = sanitize(entry.get())
            self.tree.set(row_id, "candidate", new)
            idx = self.tree.index(row_id)
            if 0 <= idx < len(self.scanned_rows):
                self.scanned_rows[idx]["candidate"] = new
            entry.destroy()

        entry.bind("<Return>", save)
        entry.bind("<FocusOut>", save)

    # ---------------- Download ----------------
    def download(self):
        dest = self.dest_var.get()
        if not dest:
            messagebox.showwarning("Pick destination",
                                   "Choose a destination folder first.")
            return
        if not self.scanned_rows:
            messagebox.showinfo("Nothing to download", "Scan a folder first.")
            return
        os.makedirs(dest, exist_ok=True)

        for i, row_id in enumerate(self.tree.get_children()):
            if i < len(self.scanned_rows):
                self.scanned_rows[i]["candidate"] = self.tree.set(row_id, "candidate")

        done_ids = self._load_log(dest)
        saved, errors = 0, 0
        for row in self.scanned_rows:
            try:
                base, ext = os.path.splitext(row["attachment"])
                cand = sanitize(row["candidate"])
                target = os.path.join(dest, f"{cand} - {sanitize(base)}{ext}")
                n = 2
                final = target
                while os.path.exists(final):
                    final = os.path.join(
                        dest, f"{cand} - {sanitize(base)} ({n}){ext}")
                    n += 1
                with open(final, "wb") as f:
                    f.write(row["data"])
                done_ids.add(f"{row['folder']}:{row['uid']}")
                saved += 1
            except Exception:
                errors += 1

        self._save_log(dest, done_ids)
        self.status.set(f"Done — {saved} file(s) saved to {dest}"
                        + (f", {errors} error(s)." if errors else "."))
        messagebox.showinfo(
            "Download complete",
            f"Saved {saved} attachment(s) to:\n{dest}"
            + (f"\n\n{errors} attachment(s) failed." if errors else ""))
        self.tree.delete(*self.tree.get_children())
        self.scanned_rows = []

    # ---------------- Log ----------------
    def _load_log(self, dest):
        p = os.path.join(dest, LOG_FILENAME)
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    return set(json.load(f))
            except Exception:
                return set()
        return set()

    def _save_log(self, dest, ids):
        try:
            with open(os.path.join(dest, LOG_FILENAME), "w", encoding="utf-8") as f:
                json.dump(sorted(ids), f)
        except Exception:
            pass


if __name__ == "__main__":
    App().mainloop()
