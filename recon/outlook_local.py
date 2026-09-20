"""Read the shared folder through the Outlook desktop client already signed in on this PC.

Uses Outlook's COM automation (pywin32). No new credentials, no app registration:
the tool sees exactly what you can already see in Outlook, nothing more.
Windows only. Same interface as GraphMail so the pipeline can swap backends.
"""
from __future__ import annotations
import re, tempfile
from datetime import datetime, timedelta
from pathlib import Path
from .graph_mail import Email, _strip_html

class OutlookMail:
    def __init__(self, folder_name: str):
        import win32com.client  # pywin32; import here so non-Windows machines can still import the package
        self.folder_name = folder_name
        self._ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self._folder = None

    def login(self, prompt=print) -> None:
        prompt("Using the Outlook session already signed in on this PC.")
        self.folder_id()

    def folder_id(self):
        if self._folder is not None:
            return self._folder
        target = self.folder_name.strip().lower()
        def walk(folders, depth=0):
            for f in folders:
                if f.Name.strip().lower() == target:
                    return f
                if depth < 4:
                    hit = walk(f.Folders, depth + 1)
                    if hit: return hit
            return None
        for store in self._ns.Folders:               # every mailbox/store shown in Outlook
            hit = walk(store.Folders)
            if hit:
                self._folder = hit; return hit
        raise LookupError(f"Folder '{self.folder_name}' not found in Outlook")

    def search(self, term: str, window_days: int, top: int = 5) -> list[dict]:
        since = (datetime.now() - timedelta(days=window_days)).strftime("%m/%d/%Y %H:%M %p")
        t = term.replace("'", "''")
        dasl = ("@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{t}%' "
                "OR \"urn:schemas:httpmail:textdescription\" LIKE '%{t}%') "
                "AND \"urn:schemas:httpmail:datereceived\" >= '{since}'").format(t=t, since=since)
        items = self.folder_id().Items
        items.Sort("[ReceivedTime]", True)
        hits = items.Restrict(dasl)
        out = []
        for it in hits:
            if getattr(it, "Class", None) != 43:     # 43 = olMail
                continue
            out.append({"id": it.EntryID, "subject": it.Subject, "receivedDateTime": it.ReceivedTime.isoformat()})
            if len(out) >= top: break
        return out

    def fetch(self, message_id: str, with_attachments=True) -> Email:
        m = self._ns.GetItemFromID(message_id)
        atts = []
        if with_attachments and m.Attachments.Count:
            tmp = Path(tempfile.mkdtemp())
            for a in m.Attachments:
                name = a.FileName or ""
                if name.lower().endswith((".pdf", ".txt", ".csv")):
                    p = tmp / re.sub(r"[^\w.\-]", "_", name)
                    a.SaveAsFile(str(p))
                    atts.append((name, p.read_bytes()))
                    p.unlink(missing_ok=True)
        sender = getattr(m, "SenderEmailAddress", "") or ""
        try:
            if m.SenderEmailType == "EX" and m.Sender:
                sender = m.Sender.GetExchangeUser().PrimarySmtpAddress or sender
        except Exception:
            pass
        body = m.Body if m.Body else _strip_html(m.HTMLBody or "")
        return Email(id=m.EntryID, subject=m.Subject or "", sender=sender,
                     received=m.ReceivedTime.isoformat(), body_text=body,
                     attachments=atts, web_link=f"outlook:{m.EntryID}")
