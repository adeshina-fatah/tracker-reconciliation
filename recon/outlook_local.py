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
        self.folder_names = [f.strip() for f in folder_name.split(",") if f.strip()]
        self.folder_name = self.folder_names[0]
        self._ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self._folders = None

    def login(self, prompt=print) -> None:
        prompt("Using the Outlook session already signed in on this PC.")
        for f in self.folders():
            try:
                n = f.Items.Count
            except Exception:
                n = "?"
            prompt(f"Searching folder: {f.FolderPath} ({n} items)")

    @staticmethod
    def _children(folder):
        """Sub-folders of a store or folder; empty if Outlook refuses to open it."""
        try:
            return list(folder.Folders)
        except Exception:
            return []

    def folders(self):
        if self._folders is None:
            self._folders = [self._find(name) for name in self.folder_names]
        return self._folders

    def folder_id(self):
        return self.folders()[0]

    def _find(self, folder_name: str):
        """Locate one folder by name, or by path like 'Mailbox Name/Inbox/Shared'.

        Stores that Outlook cannot open (offline shared mailboxes, archives,
        public folders) are skipped instead of aborting the run.
        """
        if folder_name.strip().lower() == "inbox" and len(folder_name.split("/")) == 1:
            return self._ns.GetDefaultFolder(6)   # 6 = olFolderInbox of the primary mailbox
        parts = [p.strip().lower() for p in folder_name.replace("\\", "/").split("/") if p.strip()]
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
                f"Folder '{folder_name}' not found in Outlook. "
                f"Folders visible: {sorted(set(seen))[:60]}")
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
        full = ("@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{t}%' "
                "OR \"urn:schemas:httpmail:textdescription\" LIKE '%{t}%') "
                "AND \"urn:schemas:httpmail:datereceived\" >= '{since}'").format(t=t, since=since)
        subject_only = ("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{t}%' "
                        "AND \"urn:schemas:httpmail:datereceived\" >= '{since}'").format(t=t, since=since)
        out = []
        for folder in self.folders():
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            try:
                hits = items.Restrict(full)
            except Exception:
                hits = items.Restrict(subject_only)   # some stores reject body search
            n = 0
            for it in hits:
                if getattr(it, "Class", None) != 43:     # 43 = olMail
                    continue
                out.append({"id": it.EntryID, "subject": it.Subject,
                            "receivedDateTime": it.ReceivedTime.isoformat(), "folder": folder.FolderPath})
                n += 1
                if n >= top: break
        out.sort(key=lambda m: m["receivedDateTime"], reverse=True)
        return out[:top]

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
