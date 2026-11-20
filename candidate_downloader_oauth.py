"""
Candidate Attachment Downloader — Mac (Google Sign-In version)
--------------------------------------------------------------
Downloads candidate profile attachments from a chosen Gmail folder/label
and saves them as:  <Candidate Name> - <original filename>.<ext>

Authentication:
  Uses official Google OAuth. The user clicks "Sign in with Google",
  authenticates in their own browser on Google's own page, and grants
  read-only access to Gmail. The app NEVER sees or stores their password.
  A revocable token is cached locally so sign-in is only needed rarely.

Requires a client_secret.json (OAuth Desktop client) beside the app.
"""

import os
import re
import sys
import json
import base64
import threading
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
except ImportError:
    raise SystemExit(
        "Google libraries missing. Run:\n"
        "  pip install google-api-python-client google-auth-oauthlib"
    )

# Read-only Gmail access — the narrowest scope that can read attachments.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

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

def _app_support_dir() -> str:
    """Per-OS location for the cached OAuth token.

    Windows: %APPDATA%\\CandidateDownloader
    macOS:   ~/Library/Application Support/CandidateDownloader
    other:   ~/.config/CandidateDownloader
    """
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "CandidateDownloader")
    if sys.platform == "darwin":
        return os.path.expanduser(
            "~/Library/Application Support/CandidateDownloader")
    return os.path.expanduser("~/.config/CandidateDownloader")


APP_SUPPORT = _app_support_dir()
TOKEN_PATH = os.path.join(APP_SUPPORT, "token.json")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def resource_path(name: str) -> str:
    """Find a bundled file whether running as script or as a PyInstaller app."""
    candidates = []
    if hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, name))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    candidates += [
        os.path.join(exe_dir, name),
        # macOS: alongside the .app bundle itself (Contents/MacOS/ -> ../../..)
        os.path.abspath(os.path.join(exe_dir, "..", "..", "..", name)),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), name),
        os.path.join(os.getcwd(), name),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return ""


def sanitize(name: str) -> str:
    name = INVALID_FS_CHARS.sub("", name).strip().strip(".")
    return re.sub(r"\s+", " ", name) or "Unknown"


def guess_candidate_name(subject, sender_name, sender_email):
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


def parse_from(value):
    """Split a From header into (display name, email)."""
    if not value:
        return "", ""
    m = re.match(r'\s*"?([^"<]*)"?\s*<([^>]+)>', value)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", value.strip()


def walk_parts(part):
    yield part
    for sub in part.get("parts", []) or []:
        yield from walk_parts(sub)


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Candidate Attachment Downloader")
        self.geometry("960x700")
        self.minsize(800, 580)

        self.service = None
        self.account_email = ""
        self.labels = {}          # display name -> label id
        self.scanned_rows = []

        self._build_ui()
        self.after(300, self._try_silent_signin)

    # ---------------- UI ----------------
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        acct = ttk.LabelFrame(top, text="Google account", padding=8)
        acct.pack(fill="x")
        self.acct_var = tk.StringVar(value="Not signed in")
        ttk.Label(acct, textvariable=self.acct_var).pack(side="left")
        ttk.Button(acct, text="Sign out",
                   command=self.sign_out).pack(side="right", padx=4)
        self.signin_btn = ttk.Button(acct, text="Sign in with Google",
                                     command=self.sign_in)
        self.signin_btn.pack(side="right")

        ttk.Label(top, text="Mail folder / label (the folder you see in Outlook):").pack(
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
        ttk.Button(dest_row, text="Browse…",
                   command=self._pick_dest).pack(side="left", padx=6)

        opts = ttk.Frame(top)
        opts.pack(fill="x")
        self.skip_done_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Skip emails already downloaded previously",
                        variable=self.skip_done_var).pack(side="left")
        ttk.Label(opts, text="   Scan the last").pack(side="left")
        self.days_var = tk.StringVar(value="90")
        ttk.Entry(opts, textvariable=self.days_var, width=5).pack(side="left", padx=3)
        ttk.Label(opts, text="days").pack(side="left")

        btns = ttk.Frame(top)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="1. Scan folder", command=self.scan).pack(side="left")
        ttk.Button(btns, text="2. Download attachments",
                   command=self.download).pack(side="left", padx=8)

        mid = ttk.Frame(self, padding=(10, 4, 10, 0))
        mid.pack(fill="both", expand=True)
        ttk.Label(mid, text="Preview  (double-click a Candidate Name to edit it):").pack(
            anchor="w")
        cols = ("candidate", "attachment", "subject", "received")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings")
        for cid, text, width in [
            ("candidate", "Candidate Name (editable)", 220),
            ("attachment", "Attachment", 200),
            ("subject", "Email Subject", 290),
            ("received", "Received", 120),
        ]:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=width, anchor="w")
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        self.tree.bind("<Double-1>", self._edit_cell)

        self.status = tk.StringVar(
            value="Click 'Sign in with Google' to begin.")
        ttk.Label(self, textvariable=self.status, padding=8,
                  relief="sunken", anchor="w").pack(fill="x", side="bottom")

    # ---------------- Auth ----------------
    def _try_silent_signin(self):
        """Reuse a cached token if one exists, without prompting."""
        if not os.path.exists(TOKEN_PATH):
            return
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                self._save_token(creds)
            if creds and creds.valid:
                self._activate(creds)
        except Exception:
            pass  # stale token — user just signs in again

    def sign_in(self):
        cs = resource_path("client_secret.json")
        if not cs:
            messagebox.showerror(
                "Setup file missing",
                "client_secret.json was not found.\n\n"
                "This file identifies the app to Google and must sit next to "
                "the application. Please contact whoever gave you this app.")
            return
        self.status.set("Opening your browser to sign in to Google…")
        self.update_idletasks()
        try:
            flow = InstalledAppFlow.from_client_secrets_file(cs, SCOPES)
            creds = flow.run_local_server(
                port=0,
                prompt="consent",
                access_type="offline",
                authorization_prompt_message="",
                success_message="Signed in. You can close this tab and return "
                                "to the Candidate Downloader app.")
            self._save_token(creds)
            self._activate(creds)
        except Exception as e:
            messagebox.showerror("Sign-in failed", str(e))
            self.status.set("Sign-in failed.")

    def _save_token(self, creds):
        os.makedirs(APP_SUPPORT, exist_ok=True)
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
        try:
            os.chmod(TOKEN_PATH, 0o600)
        except Exception:
            pass

    def _activate(self, creds):
        self.service = build("gmail", "v1", credentials=creds,
                             cache_discovery=False)
        prof = self.service.users().getProfile(userId="me").execute()
        self.account_email = prof.get("emailAddress", "")
        self.acct_var.set(f"Signed in as {self.account_email}")
        self.signin_btn.config(text="Switch account")
        self._load_labels()

    def sign_out(self):
        try:
            if os.path.exists(TOKEN_PATH):
                os.remove(TOKEN_PATH)
        except Exception:
            pass
        self.service = None
        self.account_email = ""
        self.acct_var.set("Not signed in")
        self.signin_btn.config(text="Sign in with Google")
        self.folder_combo["values"] = []
        self.folder_var.set("")
        self.status.set("Signed out.")

    def _load_labels(self):
        try:
            res = self.service.users().labels().list(userId="me").execute()
            labels = {}
            for lb in res.get("labels", []):
                name = lb["name"]
                if name in ("CHAT",):
                    continue
                labels[name] = lb["id"]
            self.labels = labels
            self.folder_combo["values"] = sorted(labels.keys())
            self.status.set(f"{len(labels)} folders found. "
                            "Choose the candidate folder, then Scan.")
        except Exception as e:
            self.status.set(f"Could not load folders: {e}")

    def _pick_dest(self):
        d = filedialog.askdirectory(title="Choose destination folder")
        if d:
            self.dest_var.set(d)

    # ---------------- Scan ----------------
    def scan(self):
        if not self.service:
            messagebox.showwarning("Not signed in", "Sign in with Google first.")
            return
        if not self.folder_var.get():
            messagebox.showwarning("Pick a folder", "Choose a mail folder first.")
            return
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        label_name = self.folder_var.get()
        label_id = self.labels[label_name]
        try:
            days = max(1, int(self.days_var.get()))
        except ValueError:
            days = 90
        after = (datetime.now() - timedelta(days=days)).strftime("%Y/%m/%d")

        done = set()
        if self.skip_done_var.get() and self.dest_var.get():
            done = self._load_log(self.dest_var.get())

        def ui(fn, *a):
            self.after(0, fn, *a)

        ui(self.status.set, "Scanning…")
        ui(lambda: self.tree.delete(*self.tree.get_children()))
        self.scanned_rows = []

        try:
            msgs, token = [], None
            while True:
                res = self.service.users().messages().list(
                    userId="me", labelIds=[label_id],
                    q=f"has:attachment after:{after}",
                    pageToken=token, maxResults=200).execute()
                msgs += res.get("messages", [])
                token = res.get("nextPageToken")
                if not token:
                    break

            msgs = [m for m in msgs if m["id"] not in done]
            found = 0
            for i, m in enumerate(msgs):
                ui(self.status.set,
                   f"Scanning… {i + 1}/{len(msgs)} emails, {found} attachment(s) found")
                full = self.service.users().messages().get(
                    userId="me", id=m["id"], format="full").execute()
                headers = {h["name"].lower(): h["value"]
                           for h in full.get("payload", {}).get("headers", [])}
                subject = headers.get("subject", "")
                sname, semail = parse_from(headers.get("from", ""))
                try:
                    ts = int(full.get("internalDate", "0")) / 1000
                    received = datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
                except Exception:
                    received = ""
                cand = guess_candidate_name(subject, sname, semail)

                for part in walk_parts(full.get("payload", {})):
                    fname = part.get("filename") or ""
                    if not fname:
                        continue
                    ext = os.path.splitext(fname)[1].lower()
                    if ext not in ALLOWED_EXTS:
                        continue
                    body = part.get("body", {})
                    att_id = body.get("attachmentId")
                    if not att_id:
                        continue
                    self.scanned_rows.append({
                        "msg_id": m["id"], "att_id": att_id,
                        "candidate": cand, "attachment": fname,
                        "subject": subject, "received": received,
                    })
                    found += 1
                    ui(self._insert_row, cand, fname, subject, received)

            ui(self.status.set,
               f"Scan complete — {found} attachment(s) in the last {days} days"
               + (" (previously downloaded emails skipped)." if done else "."))
        except Exception as e:
            ui(self.status.set, f"Scan error: {e}")

    def _insert_row(self, *vals):
        self.tree.insert("", "end", values=vals)

    # ---------------- Edit name ----------------
    def _edit_cell(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id or self.tree.identify_column(event.x) != "#1":
            return
        x, y, w, h = self.tree.bbox(row_id, "#1")
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

        done = self._load_log(dest)
        saved = errors = 0
        for n, row in enumerate(self.scanned_rows):
            self.status.set(f"Downloading {n + 1}/{len(self.scanned_rows)}…")
            self.update_idletasks()
            try:
                att = self.service.users().messages().attachments().get(
                    userId="me", messageId=row["msg_id"],
                    id=row["att_id"]).execute()
                data = base64.urlsafe_b64decode(att["data"])
                base_name, ext = os.path.splitext(row["attachment"])
                cand = sanitize(row["candidate"])
                final = os.path.join(dest, f"{cand} - {sanitize(base_name)}{ext}")
                k = 2
                while os.path.exists(final):
                    final = os.path.join(
                        dest, f"{cand} - {sanitize(base_name)} ({k}){ext}")
                    k += 1
                with open(final, "wb") as f:
                    f.write(data)
                done.add(row["msg_id"])
                saved += 1
            except Exception:
                errors += 1

        self._save_log(dest, done)
        self.status.set(f"Done — {saved} file(s) saved"
                        + (f", {errors} error(s)." if errors else "."))
        messagebox.showinfo("Download complete",
                            f"Saved {saved} attachment(s) to:\n{dest}"
                            + (f"\n\n{errors} failed." if errors else ""))
        self.tree.delete(*self.tree.get_children())
        self.scanned_rows = []

    # ---------------- Log ----------------
    def _load_log(self, dest):
        p = os.path.join(dest, LOG_FILENAME)
        if os.path.exists(p):
            try:
                with open(p, encoding="utf-8") as f:
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
