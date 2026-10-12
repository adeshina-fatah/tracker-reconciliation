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
    # specific to the charge; generic words removed so ordinary mail does not match
    "stock count", "stock take", "stocktake", "cycle count", "stock discrepanc", "stock variance", "variance report",
    "write-off", "write off", "service exchange", "SX form", "SX store", "SX unit", "drifter part", "drifter stock",
    "stock adjustment", "issued to job", "returned to stock", "reconciliation", "root cause", "corrective action",
    "stock control", "stock record", "physical count", "count sheet", "store access", "access to the store",
    "shelving", "shelves", "bins", "labelling", "workload", "overtime", "backlog", "capacity",
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
    ap.add_argument("--min-hits", type=int, default=2, help="distinct keywords a message must contain (default 2)")
    ap.add_argument("--max", type=int, default=0, help="stop after this many matches (0 = no limit)")
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
    out = Path(a.out).resolve(); (out / "msg").mkdir(parents=True, exist_ok=True)
    failed_saves = 0

    rows, seen, skipped = [], set(), []

    def safe_get(obj, attr, default=""):
        try:
            v = getattr(obj, attr)
            return default if v is None else v
        except Exception:
            return default

    for folder in folders:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        restrict = ("[ReceivedTime] >= '{}' AND [ReceivedTime] <= '{}'"
                    .format(since.strftime("%m/%d/%Y 00:00"), until.strftime("%m/%d/%Y 23:59")))
        print(f"Scanning {folder.FolderPath} ...")
        n = 0
        for it in items.Restrict(restrict):
            if safe_get(it, "Class", None) != 43:
                continue
            n += 1
            if n % 500 == 0:
                print(f"  ... {n} scanned, {len(rows)} matched")
            subject = safe_get(it, "Subject")
            try:
                body = it.Body or ""
            except Exception as e:          # sync-conflict or corrupt item: Outlook refuses the body
                body = ""
                skipped.append((safe_get(it, "ReceivedTime"), subject, str(e)[:80]))
            text = f"{subject}\n{body}"
            hits = sorted({m.group(0).lower() for m in rx.finditer(text)})
            entry_id = safe_get(it, "EntryID", None)
            if len(hits) < a.min_hits or not entry_id or entry_id in seen:
                continue
            seen.add(entry_id)
            when = safe_get(it, "ReceivedTime", None)
            if when is None:
                continue
            stamp = when.strftime("%Y-%m-%d_%H%M")
            safe = re.sub(r"[^A-Za-z0-9\- ]", "_", (it.Subject or "no subject")).strip(" ._")[:40] or "no_subject"
            fname = f"{stamp} {safe}.msg"
            if not a.no_msg:
                target = out / "msg" / fname
                try:
                    it.SaveAs(str(target), 3)            # 3 = olMSG; needs an absolute path
                except Exception:
                    # Outlook refused (rights-managed mail, odd item, policy): keep the text instead
                    fname = fname[:-4] + ".txt"
                    try:
                        (out / "msg" / fname).write_text(
                            f"From: {safe_get(it, 'SenderName')}\nTo: {safe_get(it, 'To')}\nCC: {safe_get(it, 'CC')}\n"
                            f"Date: {when}\nSubject: {subject}\n\n{body}",
                            encoding="utf-8")
                    except Exception:
                        fname = "(could not save)"
                    failed_saves += 1
            body = re.sub(r"\s+", " ", body)
            first = rx.search(body)
            snippet = body[max(0, first.start() - 150): first.start() + 250] if first else body[:300]
            try:
                att_count = it.Attachments.Count
            except Exception:
                att_count = ""
            rows.append([when.strftime("%Y-%m-%d %H:%M"), folder.Name, safe_get(it, "SenderName"),
                         safe_get(it, "To"), safe_get(it, "CC"), subject, ", ".join(hits), att_count,
                         snippet, fname if not a.no_msg else ""])
        print(f"  {n} emails in window, {len(rows)} matches so far")
        if a.max and len(rows) >= a.max:
            print("  --max reached, stopping"); break

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
    if skipped:
        log = out / "skipped-items.txt"
        log.write_text("\n".join(f"{d} | {s} | {e}" for d, s, e in skipped), encoding="utf-8")
        print(f"{len(skipped)} item(s) could not be read (Outlook sync conflicts); subjects listed in {log}. "
              "Open them in Outlook to resolve, then re-run if any look relevant.")
    if failed_saves:
        print(f"{failed_saves} message(s) could not be saved as .msg and were saved as .txt instead")

if __name__ == "__main__":
    main()
