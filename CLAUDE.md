# Candidate Downloader

A cross-platform (macOS + Windows) desktop app that downloads LinkedIn candidate
CVs (and other profile attachments) from Gmail and saves them with
human-readable names of the form:

```
Candidate Name - originalfilename.ext
```

## What it does

The user signs in to their own Google account via official Google OAuth
("Sign in with Google"). The app requests **read-only** Gmail access
(`gmail.readonly`) — it never sees or stores the user's password. It then:

1. Lists the user's Gmail labels/folders so they can pick the folder where
   LinkedIn candidate emails land.
2. Scans that folder (within a configurable look-back window, default 90 days)
   for messages with attachments.
3. Infers each candidate's name from the email subject/sender (see
   `SUBJECT_PATTERNS` in the source) and shows an editable preview table.
4. Downloads the allowed attachment types (`.pdf`, `.doc`, `.docx`, `.rtf`,
   `.odt`) into a chosen destination folder, named
   `<Candidate Name> - <original filename>.<ext>`.

A `_downloaded_log.json` is kept in the destination folder so already-downloaded
emails can be skipped on re-runs.

## Layout

- `candidate_downloader_oauth.py` — **the current app** (Tkinter GUI, Google
  OAuth). This is the file that gets built and shipped.
- `candidate_attachment_downloader_mac.py` — older variant, kept for reference.
- `client_secret.json` — Google OAuth **Desktop** client credentials.
  **Not committed** (git-ignored). You must place this file in the **project
  root** before building — the build bundles it into the app. See "Obtaining
  client_secret.json" below.
- `Candidate Downloader.spec` — PyInstaller spec produced by the build.
- `build_oauth.sh` — the canonical build script (see below).
- `build.sh` — older build script, kept for reference.

## Obtaining client_secret.json

`client_secret.json` is **not** in the repo — it is a secret and must be
supplied locally. Create it once in the Google Cloud Console and place it in the
**project root** (next to `candidate_downloader_oauth.py`):

1. https://console.cloud.google.com → create/select a project.
2. Enable the **Gmail API**.
3. Configure the **OAuth consent screen**.
4. **Credentials → Create OAuth client ID → Application type: Desktop app**.
5. Download the JSON, rename it to `client_secret.json`, and drop it in the
   project root.

At runtime the app also accepts the file sitting next to the `.app` bundle, so a
recipient can replace it without a rebuild.

## Runtime auth details

- Scope: `https://www.googleapis.com/auth/gmail.readonly`.
- The OAuth token is cached per-OS (see `_app_support_dir()` in the source) as
  `token.json` (mode 600 on POSIX) so sign-in is only needed rarely.
  "Sign out" deletes it. Locations:
  - **Windows:** `%APPDATA%\CandidateDownloader`
  - **macOS:** `~/Library/Application Support/CandidateDownloader`
  - **other:** `~/.config/CandidateDownloader`
- `client_secret.json` is looked up next to the app bundle and inside the
  PyInstaller bundle (`resource_path()` in the source).

## Building the macOS app

Requirements: a Python with `tkinter` and PyInstaller.

> **Python version note:** PyInstaller may not support the very latest CPython.
> If a build fails on the system Python, install `python@3.12` via Homebrew and
> build with that instead.

The build is a `--windowed --onedir` PyInstaller bundle named
**"Candidate Downloader"**, with `client_secret.json` bundled via `--add-data`
and `--collect-all` for `googleapiclient`, `google_auth_oauthlib` and
`google.auth`. The result is ad-hoc codesigned and zipped for delivery.

```bash
bash build_oauth.sh
```

Output:
- `dist/Candidate Downloader.app` — the app bundle
- `dist/Candidate Downloader.zip` — the zip to distribute

First-run instructions for recipients: right-click the app → **Open** → **Open**
(to bypass Gatekeeper for an ad-hoc-signed app), then click
**Sign in with Google**.

## Building the Windows app (GitHub Actions)

Windows `.exe` builds run in CI, not locally, so no Windows machine is needed.

- Workflow: `.github/workflows/build-windows.yml`, runs on `windows-latest`,
  triggered manually via **workflow_dispatch**.
- It sets up Python 3.12, installs the same deps, writes `client_secret.json`
  from the `CLIENT_SECRET_JSON` Actions secret (validating it parses as JSON),
  then builds a **`--windowed --onefile`** PyInstaller exe named
  **"Candidate Downloader"** (same `--add-data` / `--collect-all` /
  `--hidden-import` flags as the Mac build, but with the Windows `;` path
  separator in `--add-data`).
- Output: uploaded as the artifact **`Candidate-Downloader-Windows`**
  (`dist/Candidate Downloader.exe`).

The OAuth secret is stored as a repo Actions secret so it stays out of git:

```bash
gh secret set CLIENT_SECRET_JSON < client_secret.json
```

Trigger + download:

```bash
gh workflow run build-windows.yml
gh run watch <run-id>
gh run download <run-id> -n Candidate-Downloader-Windows -D windows-build
```

## Dependencies

Installed into a throwaway build venv (`.buildenv`):

- `google-api-python-client`
- `google-auth-oauthlib`
- `google-auth-httplib2`
- `pyinstaller`

## Notes

- `.venv`, `.buildenv`, `build/`, `dist/`, `windows-build/`, `*.spec`, and
  `client_secret.json` are git-ignored (build artifacts / environments,
  CI download folder, and the OAuth secret).
- The app is cross-platform (macOS + Windows). The only OS-specific code is the
  token cache location in `_app_support_dir()`; the Mac build produces a `.app`
  bundle and the Windows CI build produces a single `.exe`.
