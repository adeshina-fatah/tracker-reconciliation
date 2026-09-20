"""Streamlit UI: upload the weekly workbook, review the queue, run, download results."""
import json, tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
from recon.pipeline import load_config, build_queue, run

st.set_page_config(page_title="Tracker Reconciliation", layout="wide")
st.title("Weekly tracker reconciliation")
cfg = load_config()

up = st.file_uploader("Upload the weekly Access export (.xlsx with red/amber shaded Comments)", type="xlsx")
if not up:
    st.info("Upload the workbook your manager sends. Only red and amber shaded rows are processed.")
    st.stop()

tmp = Path(tempfile.mkdtemp()) / up.name
tmp.write_bytes(up.getvalue())
queue = build_queue(str(tmp), cfg)
df = pd.DataFrame(queue)

c1, c2, c3 = st.columns(3)
c1.metric("Rows to work", len(df))
c2.metric("Red (ETA / dispatch)", int((df["shade"] == "red").sum()))
c3.metric("Amber (awaiting PO)", int((df["shade"] == "amber").sum()))

sheet = st.multiselect("Sheets", sorted(df["sheet"].unique()), default=sorted(df["sheet"].unique()))
shade = st.multiselect("Shade", ["red", "amber"], default=["red", "amber"])
view = df[df["sheet"].isin(sheet) & df["shade"].isin(shade)]
st.dataframe(view[["sheet", "row", "shade", "job_no", "rfq", "order_ref", "customer", "status",
                   "current_comment", "flags"]], use_container_width=True, height=380)
flagged = df[df["flags"].map(bool)]
if len(flagged):
    st.warning(f"{len(flagged)} rows have data-quality flags (bad job number format, duplicates, no identifiers). "
               "They are still searched but check them by hand.")

st.divider()
col_a, col_b = st.columns([1, 3])
dry = col_a.toggle("Dry run (no mailbox, no AI)", value=True)
limit = col_a.number_input("Limit rows (0 = all)", min_value=0, value=0)
if col_a.button("Run reconciliation", type="primary"):
    log = col_b.empty(); lines = []
    def progress(msg):
        lines.append(msg); log.code("\n".join(lines[-15:]))
    with st.spinner("Working..."):
        r = run(str(tmp), cfg, dry_run=dry, limit=limit or None, progress=progress)
    st.success(f"Done. {len(r['results'])} rows processed.")
    res = pd.DataFrame(r["results"])
    show = [c for c in ["sheet", "row", "shade", "job_no", "proposed_comment", "proposed_status",
                        "confidence", "evidence", "email_date", "flags"] if c in res.columns]
    st.dataframe(res[show], use_container_width=True)
    st.download_button("Download reconciled workbook", Path(r["workbook"]).read_bytes(),
                       file_name=Path(r["workbook"]).name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("The original workbook is never modified. Review proposals, then paste approved comments into the tracker.")
