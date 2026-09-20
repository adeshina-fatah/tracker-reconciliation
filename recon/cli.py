import argparse, logging
from .pipeline import load_config, run

def main():
    ap = argparse.ArgumentParser(description="Reconcile shaded tracker rows against the shared mailbox")
    ap.add_argument("xlsx"); ap.add_argument("--config", default="config.yaml")
    ap.add_argument("--dry-run", action="store_true", help="extract the work queue only; no mailbox, no model")
    ap.add_argument("--limit", type=int); ap.add_argument("--out", default="output")
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    r = run(a.xlsx, load_config(a.config), a.out, dry_run=a.dry_run, limit=a.limit, progress=print)
    reds = sum(q["shade"] == "red" for q in r["queue"]); ambers = len(r["queue"]) - reds
    print(f"\n{len(r['queue'])} rows queued ({reds} red, {ambers} amber). Output: {r['workbook']}")

if __name__ == "__main__":
    main()
