# Tracker reconciliation

Reads the weekly Access export, picks out the red and amber shaded Comments cells,
searches one shared folder in your Outlook mailbox for each job, and proposes an
updated comment with the evidence. Output is a copy of the workbook with extra
columns; the original is never modified.

## Setup
```
python -m venv .venv && . .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env                              # fill in values
```
Run this on the Windows PC where Outlook is signed in. The default backend reads the
shared folder through Outlook itself, so no extra credentials are needed: set
`MAIL_FOLDER_NAME` to the folder's display name and you are done. Outlook must be open.

Optional `MAIL_BACKEND=graph` uses Microsoft Graph instead and needs an Entra app
registration (delegated `Mail.Read`) from IT.

## Use
Offline check of the queue (no mailbox, no AI):
```
python -m recon.cli "path/to/report.xlsx" --dry-run
```
Full run, or the UI:
```
python -m recon.cli "path/to/report.xlsx"
streamlit run app.py
```
With the Graph backend, the first run prompts a device-code sign-in; the token is cached
locally in `.token_cache.json` (gitignored).

## Data handling
Only shaded rows are processed. Only the matched emails and PDF/text attachments
are sent to the model, with account-number-like strings redacted. `output/audit.log`
records each run. Delete `output/` per your retention policy.
