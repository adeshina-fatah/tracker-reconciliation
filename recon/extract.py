"""Turn matched emails into a structured reconciliation proposal with Claude."""
from __future__ import annotations
import io, re
from typing import Literal
from pydantic import BaseModel, Field
import anthropic
from pypdf import PdfReader

class Proposal(BaseModel):
    proposed_comment: str = Field(description="One-line comment in the tracker's house style, e.g. 'ETA 30.09.26', 'Received customer PO 4500600484 24.08.26', 'Dispatched 24.08.26'. Empty string if no evidence.")
    proposed_status: str | None = Field(description="Suggested Job Status value if the evidence shows it should change, else null.")
    po_number: str | None
    po_date: str | None = Field(description="ISO date or null")
    eta_date: str | None
    dispatch_date: str | None
    invoice_number: str | None
    evidence: str = Field(description="Short verbatim quote from the email or attachment that supports the proposal.")
    confidence: Literal["high", "medium", "low"]
    conflict: str | None = Field(description="Describe if emails disagree, or the job number in the email does not match, else null.")

SYSTEM = """You reconcile a mining-equipment repair job tracker against emails from a shared mailbox.
You are given one tracker row (identifiers, current status, current comment, and what the reviewer needs)
and the text of the most recent matching emails and attachments. Fill the schema strictly from the evidence.
Never invent dates or numbers. If the emails do not answer the objective, leave fields null,
set confidence to low and say so in evidence. Quote evidence verbatim. Dates in the comment use DD.MM.YY."""

REDACT = [
    (re.compile(r"\b\d{2}[- ]?\d{2}[- ]?\d{2}[- ]?\d{2}[- ]?\d{2}\b"), "[redacted-account]"),  # bank-ish numbers
    (re.compile(r"\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b"), "[redacted-iban]"),
]

def pdf_text(data: bytes, max_pages=6) -> str:
    try:
        r = PdfReader(io.BytesIO(data))
        return "\n".join((p.extract_text() or "") for p in r.pages[:max_pages])
    except Exception as e:  # scanned PDFs land here; OCR is a later add-on
        return f"[could not extract PDF text: {e}]"

def redact(text: str) -> str:
    for rx, rep in REDACT:
        text = rx.sub(rep, text)
    return text

def build_context(item: dict, emails: list) -> str:
    parts = [f"TRACKER ROW\nsheet={item['sheet']} row={item['row']} shade={item['shade']}",
             f"job_no={item.get('job_no')} rfq={item.get('rfq')} order_ref={item.get('order_ref')} "
             f"part_no={item.get('part_no')} customer={item.get('customer')}",
             f"current_status={item.get('status')}\ncurrent_comment={item.get('current_comment')}",
             f"OBJECTIVE: {item['objective']}\n"]
    for i, e in enumerate(emails, 1):
        parts.append(f"=== EMAIL {i} | {e.received} | from {e.sender} | subject: {e.subject}\n{redact(e.body_text[:6000])}")
        for name, data in e.attachments:
            txt = pdf_text(data) if name.lower().endswith(".pdf") else data.decode("utf-8", "ignore")
            parts.append(f"--- ATTACHMENT {name}\n{redact(txt[:6000])}")
    return "\n\n".join(parts)

class Extractor:
    def __init__(self, model: str = "claude-opus-5"):
        self.client = anthropic.Anthropic()   # reads ANTHROPIC_API_KEY
        self.model = model

    def propose(self, item: dict, emails: list) -> Proposal:
        resp = self.client.messages.parse(
            model=self.model, max_tokens=4000, system=SYSTEM,
            messages=[{"role": "user", "content": build_context(item, emails)}],
            output_format=Proposal)
        return resp.parsed_output
