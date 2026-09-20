"""Orchestrates: workbook -> shaded rows -> mailbox search -> extraction -> result workbook."""
from __future__ import annotations
import json, logging, os
from datetime import datetime
from pathlib import Path
import yaml
from dotenv import load_dotenv
from .workbook import load_work_items, items_to_records, write_results
from .identifiers import search_terms

log = logging.getLogger("recon")

def load_config(path="config.yaml") -> dict:
    return yaml.safe_load(Path(path).read_text())

def make_mailbox():
    """MAIL_BACKEND=outlook (default, uses the signed-in Outlook desktop client) or graph."""
    backend = os.environ.get("MAIL_BACKEND", "outlook").lower()
    folder = os.environ.get("MAIL_FOLDER_NAME") or os.environ.get("GRAPH_FOLDER_NAME")
    if not folder:
        raise RuntimeError("Set MAIL_FOLDER_NAME in .env to the shared folder's display name")
    if backend == "graph":
        from .graph_mail import GraphMail
        return GraphMail(os.environ["GRAPH_CLIENT_ID"], os.environ["GRAPH_TENANT_ID"], folder)
    from .outlook_local import OutlookMail
    return OutlookMail(folder)

def build_queue(xlsx: str, cfg: dict) -> list[dict]:
    recs = items_to_records(load_work_items(xlsx, cfg))
    for r in recs:
        r["search_terms"] = search_terms(r)
    return recs

def run(xlsx: str, cfg: dict, out_dir="output", dry_run=False, progress=None, limit=None) -> dict:
    """dry_run: build the queue and write it out without touching the mailbox or the model."""
    load_dotenv()
    stamp = datetime.now().strftime("%Y%m%d-%H%M")
    out = Path(out_dir); out.mkdir(exist_ok=True)
    queue = build_queue(xlsx, cfg)
    if limit: queue = queue[:limit]
    (out / f"queue-{stamp}.json").write_text(json.dumps(queue, indent=2, default=str))
    results = []
    if dry_run:
        for q in queue:
            results.append({**q, "proposed_comment": None, "confidence": None,
                            "evidence": "dry run: mailbox not queried"})
    else:
        from .extract import Extractor
        gm = make_mailbox()
        gm.login(prompt=progress or print)
        ex = Extractor(cfg.get("model", "claude-opus-5"))
        for n, q in enumerate(queue, 1):
            if progress: progress(f"[{n}/{len(queue)}] job {q['job_no']} ({q['shade']})")
            res = dict(q)
            try:
                msgs = []
                for _, term in q["search_terms"]:
                    msgs = gm.search(term, cfg["search_window_days"], cfg["max_emails_per_row"])
                    if msgs: break
                if not msgs:
                    res.update(proposed_comment="", confidence="low", evidence="No matching emails in window")
                else:
                    emails = [gm.fetch(m["id"]) for m in msgs]
                    p = ex.propose(q, emails)
                    res.update(p.model_dump(), email_date=emails[0].received[:10],
                               source=f"{emails[0].subject} / {emails[0].sender} / {emails[0].web_link}")
                    if p.conflict: res["flags"] = res.get("flags", []) + ["conflict"]
            except Exception as e:
                log.exception("row failed"); res.update(confidence="low", evidence=f"error: {e}")
                res["flags"] = res.get("flags", []) + ["error"]
            results.append(res)
    result_path = write_results(xlsx, out / f"reconciled-{stamp}.xlsx", results, cfg)
    (out / f"results-{stamp}.json").write_text(json.dumps(results, indent=2, default=str))
    # audit line: who ran what, when, on which file
    with (out / "audit.log").open("a") as f:
        f.write(f"{datetime.now().isoformat()} user={os.environ.get('USERNAME') or os.environ.get('USER')} "
                f"file={xlsx} rows={len(queue)} dry_run={dry_run} out={result_path}\n")
    return {"queue": queue, "results": results, "workbook": str(result_path)}
