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
        import pythoncom, win32com.client  # pywin32; imported here so non-Windows machines can import the package
        pythoncom.CoInitialize()          # Streamlit runs us in a worker thread; COM needs this per thread
        self.folder_name = folder_name
        self._ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self._folder = None

    def login(self, prompt=print) -> None:
        prompt("Using the Outlook session already signed in on this PC.")
        f = self.folder_id()
        prompt(f"Found folder: {f.FolderPath}")

    @staticmethod
    def _children(folder):
        """Sub-folders of a store or folder; empty if Outlook refuses to open it."""
        try:
            return list(folder.Folders)
        except Exception:
            return []

    def folder_id(self):
        """Locate the folder by name, or by path like 'Mailbox Name/Inbox/Shared'.

        Stores that Outlook cannot open (offline shared mailboxes, archives,
        public folders) are skipped instead of aborting the run.
        """
        if self._folder is not None:
            return self._folder
        parts = [p.strip().lower() for p in self.folder_name.replace("\\", "/").split("/") if p.strip()]
        target = parts[-1]
        seen = []

        def walk(folders, depth=0):
            for f in folders:
                try:
                    name = f.Name
                except Exception:
                    continue
                seen.append(name)
                if name.strip().lower() == target:
                    if len(parts) == 1 or self._path_matches(f, parts):
                        return f
                if depth < 5:
                    hit = walk(self._children(f), depth + 1)
                    if hit:
                        return hit
            return None

        hit = walk(self._children(self._ns))
        if hit is None:
            raise LookupError(
                f"Folder '{self.folder_name}' not found in Outlook. "
                f"Folders visible: {sorted(set(seen))[:60]}")
        self._folder = hit
        return hit

    @staticmethod
    def _path_matches(folder, parts) -> bool:
        try:
            path = [p.lower() for p in folder.FolderPath.strip("\\").split("\\")]
        except Exception:
            return False
        return path[-len(parts):] == parts

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
