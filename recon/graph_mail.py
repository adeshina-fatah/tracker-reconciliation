"""Microsoft Graph access to one folder of the signed-in user's mailbox.

Delegated auth (device code) with Mail.Read only. Token cache is a local file
the user owns; no client secret is ever stored. Searches are scoped to the
configured folder and date window so we never read more mail than needed.
"""
from __future__ import annotations
import base64, json, os
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
import msal, requests

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read"]
CACHE_FILE = Path(".token_cache.json")

@dataclass
class Email:
    id: str
    subject: str
    sender: str
    received: str
    body_text: str
    attachments: list[tuple[str, bytes]]   # (filename, content)
    web_link: str

class GraphMail:
    def __init__(self, client_id: str, tenant_id: str, folder_name: str):
        self.folder_name = folder_name
        cache = msal.SerializableTokenCache()
        if CACHE_FILE.exists():
            cache.deserialize(CACHE_FILE.read_text())
        self._cache = cache
        self._app = msal.PublicClientApplication(
            client_id, authority=f"https://login.microsoftonline.com/{tenant_id}", token_cache=cache)
        self._token = None
        self._folder_id = None

    # ---- auth -------------------------------------------------------------
    def login(self, prompt=print) -> None:
        accounts = self._app.get_accounts()
        result = self._app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
        if not result:
            flow = self._app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"Device flow failed: {flow}")
            prompt(flow["message"])
            result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(result.get("error_description", result))
        self._token = result["access_token"]
        if self._cache.has_state_changed:
            CACHE_FILE.write_text(self._cache.serialize())
            try: os.chmod(CACHE_FILE, 0o600)
            except OSError: pass

    def _get(self, url: str, **params) -> dict:
        r = requests.get(url, headers={"Authorization": f"Bearer {self._token}",
                                       "ConsistencyLevel": "eventual"}, params=params, timeout=60)
        r.raise_for_status()
        return r.json()

    # ---- folder -----------------------------------------------------------
    def folder_id(self) -> str:
        if self._folder_id:
            return self._folder_id
        # search top level, then one level down (shared folders often sit under Inbox)
        for url in (f"{GRAPH}/me/mailFolders", f"{GRAPH}/me/mailFolders/inbox/childFolders"):
            data = self._get(url, **{"$top": 200, "$select": "id,displayName"})
            for f in data.get("value", []):
                if f["displayName"].strip().lower() == self.folder_name.strip().lower():
                    self._folder_id = f["id"]; return f["id"]
        raise LookupError(f"Folder '{self.folder_name}' not found in mailbox")

    # ---- search -----------------------------------------------------------
    def search(self, term: str, window_days: int, top: int = 5) -> list[dict]:
        since = (datetime.now(timezone.utc) - timedelta(days=window_days)).strftime("%Y-%m-%d")
        data = self._get(
            f"{GRAPH}/me/mailFolders/{self.folder_id()}/messages",
            **{"$search": f'"{term}" AND received>={since}', "$top": top,
               "$select": "id,subject,from,receivedDateTime,hasAttachments,webLink"})
        msgs = data.get("value", [])
        msgs.sort(key=lambda m: m["receivedDateTime"], reverse=True)
        return msgs

    def fetch(self, message_id: str, with_attachments=True) -> Email:
        m = self._get(f"{GRAPH}/me/messages/{message_id}",
                      **{"$select": "id,subject,from,receivedDateTime,body,webLink,hasAttachments",
                         "Prefer": "outlook.body-content-type=text"})
        atts = []
        if with_attachments and m.get("hasAttachments"):
            for a in self._get(f"{GRAPH}/me/messages/{message_id}/attachments").get("value", []):
                if a.get("@odata.type") == "#microsoft.graph.fileAttachment" and a.get("contentBytes"):
                    name = a.get("name", "")
                    if name.lower().endswith((".pdf", ".txt", ".csv")) and a.get("size", 0) < 8_000_000:
                        atts.append((name, base64.b64decode(a["contentBytes"])))
        return Email(
            id=m["id"], subject=m.get("subject") or "",
            sender=(m.get("from") or {}).get("emailAddress", {}).get("address", ""),
            received=m["receivedDateTime"], body_text=_strip_html(m["body"]["content"]),
            attachments=atts, web_link=m.get("webLink", ""))

def _strip_html(s: str) -> str:
    import re, html
    s = re.sub(r"<(script|style).*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<br\s*/?>|</p>|</div>|</tr>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"[ \t]+", " ", html.unescape(s)).strip()
