"""Collect emails from your OWN mailbox that relate to a set of topics, for a hearing or review.

Run on the Windows PC where Outlook is signed in, under your own login only.
Nothing is sent anywhere: matches are saved as .msg files plus an Excel index
in a local folder. Keep that folder on the work PC or hand it to your
representative; do not forward bulk company mail to a personal address.

Usage:
  python -m recon.evidence --since 2024-01-01 --out evidence
  python -m recon.evidence --since 2024-01-01 --keywords my_keywords.txt --folders "Inbox, Sent Items, Archive"
"""
from __future__ import annotations
import argparse, re
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font

DEFAULT_KEYWORDS = [
    # the subject matter of the charge
    "stock count", "stock take", "stocktake", "cycle count", "stock discrepanc", "variance", "write-off", "write off",
    "service exchange", "SX form", "SX store", "SX", "drifter", "adjustment", "issued to job", "returned to stock",
    "reconcil", "investigation", "root cause", "corrective", "stock control", "bin", "shelv", "labell",
    # access and control (who else could move stock)
    "access", "store access", "key", "took", "taken", "borrowed", "without",
    # workload, capacity and support requests
    "overtime", "workload", "capacity", "assist", "help", "backlog", "pressure", "deadline", "close of business",
    "laptop", "no access", "urgent", "by COB", "request",
    # people named in the matter (edit to match your colleagues)
    "Kgothatso", "Thelma",
]

def load_keywords(path: str | None) -> list[str]:
    if not path:
        return DEFAULT_KEYWORDS
    return [l.strip() for l in Path(path).read_text(encoding="utf-8").splitlines() if l.strip() and not l.startswith("#")]

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--since", required=True, help="YYYY-MM-DD, earliest email date to include")
    ap.add_argument("--until", default=None, help="YYYY-MM-DD, latest email date (default today)")
    ap.add_argument("--folders", default="Inbox, Sent Items", help="comma-separated folder names or paths")
    ap.add_argument("--keywords", default=None, help="text file, one keyword per line (default: built-in list)")
    ap.add_argument("--out", default="evidence")
    ap.add_argument("--no-msg", action="store_true", help="index only, do not save .msg copies")
    a = ap.parse_args()

    import pythoncom, win32com.client
    pythoncom.CoInitialize()
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    from .outlook_local import OutlookMail
    finder = OutlookMail.__new__(OutlookMail); finder._ns = ns; finder._folders = None
    finder.folder_names = [f.strip() for f in a.folders.split(",") if f.strip()]
    folders = []
    for name in finder.folder_names:
        if name.lower() == "sent items":
            folders.append(ns.GetDefaultFolder(5))      # olFolderSentMail
        else:
            folders.append(finder._find(name))

    kws = load_keywords(a.keywords)
    rx = re.compile("|".join(re.escape(k) for k in kws), re.I)
    since = datetime.strptime(a.since, "%Y-%m-%d")
    until = datetime.strptime(a.until, "%Y-%m-%d") if a.until else datetime.now()
    out = Path(a.out); (out / "msg").mkdir(parents=True, exist_ok=True)

    rows, seen = [], set()
    for folder in folders:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        restrict = ("[ReceivedTime] >= '{}' AND [ReceivedTime] <= '{}'"
                    .format(since.strftime("%m/%d/%Y 00:00"), until.strftime("%m/%d/%Y 23:59")))
        print(f"Scanning {folder.FolderPath} ...")
        n = 0
        for it in items.Restrict(restrict):
            if getattr(it, "Class", None) != 43:
                continue
            n += 1
            text = f"{it.Subject or ''}\n{it.Body or ''}"
            hits = sorted({m.group(0).lower() for m in rx.finditer(text)})
            if not hits or it.EntryID in seen:
                continue
            seen.add(it.EntryID)
            when = it.ReceivedTime
            stamp = when.strftime("%Y-%m-%d_%H%M")
            safe = re.sub(r"[^\w\- ]", "_", (it.Subject or "no subject"))[:60].strip()
            fname = f"{stamp} {safe}.msg"
            if not a.no_msg:
                it.SaveAs(str(out / "msg" / fname), 3)    # 3 = olMSG
            body = re.sub(r"\s+", " ", it.Body or "")
            first = rx.search(body)
            snippet = body[max(0, first.start() - 150): first.start() + 250] if first else body[:300]
            rows.append([when.strftime("%Y-%m-%d %H:%M"), folder.Name, it.SenderName, it.To, it.CC,
                         it.Subject, ", ".join(hits), it.Attachments.Count, snippet, fname if not a.no_msg else ""])
        print(f"  {n} emails in window, {len(rows)} matches so far")

    rows.sort(key=lambda r: r[0])
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Index"
    hdr = ["Date", "Folder", "From", "To", "CC", "Subject", "Matched keywords", "Attachments", "Snippet", "Saved file"]
    ws.append(hdr)
    for c in ws[1]: c.font = Font(bold=True)
    for r in rows: ws.append(r)
    for col, w in zip("ABCDEFGHIJ", (17, 12, 24, 30, 24, 50, 30, 11, 80, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    path = out / f"evidence-index-{datetime.now():%Y%m%d-%H%M}.xlsx"
    wb.save(path)
    print(f"\n{len(rows)} matching emails. Index: {path}\nSaved copies: {out / 'msg'}")

if __name__ == "__main__":
    main()
