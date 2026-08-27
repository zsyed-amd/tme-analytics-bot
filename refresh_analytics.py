#!/usr/bin/env python3
"""
TME Content Analytics — Monthly Refresh

Uses the shared m365 token from ~/.config/microsoft-graph/token.json
(same auth as m365-teams/calendar/email skills — no app registration needed).

To authenticate for the first time:
  python3 /home/zsyed/.claude/skills/m365-teams/scripts/auth.py

For Agent Hub runs, set MS_REFRESH_TOKEN env var and the client will
exchange it for a fresh access token automatically.
"""

import sys
import os
import base64
import json
import datetime
import urllib.parse
import urllib.request
import urllib.error
import ssl
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Inline token client (mirrors graph_client.py from m365-teams skill)
# Supports both local token file and MS_REFRESH_TOKEN env var for Agent Hub
# ---------------------------------------------------------------------------

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"  # Teams Desktop — pre-authorized
TENANT_ID = "3dd8961f-e488-4e60-8e11-a82d994e183d"  # AMD tenant — constant, used when no token file
SCOPES = "Files.ReadWrite.All Sites.ReadWrite.All offline_access"
# In Agent Hub sandbox the repo is cloned to /home/agent/workspace/tme-analytics-bot/
# and only that subdirectory is writable, so store the token there.
_REPO_DIR = Path(__file__).parent
_SANDBOX_TOKEN = _REPO_DIR / "token.json"
TOKEN_FILE = (
    _SANDBOX_TOKEN
    if _SANDBOX_TOKEN.exists() or Path("/home/agent/workspace").exists()
    else Path.home() / ".config" / "microsoft-graph" / "token.json"
)

LIST_SITE_ID = "amdcloud-my.sharepoint.com,fc60e54f-ec1b-469b-93c2-d79471b2a67c,cd704538-2cf6-49ae-a1a4-4343ec2e3b38"
LIST_ID = "dc498da1-bb6b-4935-aa2a-07ab08e56486"  # TME Content analytics
COLLATERAL_SITE_ID = "amdcloud.sharepoint.com,8a79450f-1df2-4861-8fe8-73ceec656271,ea2697f2-a242-4280-988a-09bdedf5919c"


def _ssl_ctx():
    for cand in (os.environ.get("SSL_CERT_FILE"), "/etc/pki/tls/certs/ca-bundle.crt",
                 "/etc/ssl/certs/ca-certificates.crt"):
        if cand and os.path.isfile(cand):
            return ssl.create_default_context(cafile=cand)
    return ssl.create_default_context()


class TokenClient:
    def __init__(self):
        self._access_token = ""
        self._refresh_token = os.environ.get("MS_REFRESH_TOKEN", "")
        # Agent Hub passes the token base64-encoded (copy-safe: no periods to corrupt)
        if not self._refresh_token:
            b64 = os.environ.get("MS_REFRESH_TOKEN_B64", "").strip()
            if b64:
                self._refresh_token = base64.b64decode(b64).decode()
        self._expires_at = 0.0
        # Env override, else constant; token file (if present) can still override below
        self._tenant_id = os.environ.get("MS_TENANT_ID", TENANT_ID)
        self._ctx = _ssl_ctx()
        self._load_file()

    def _load_file(self):
        if TOKEN_FILE.exists():
            d = json.loads(TOKEN_FILE.read_text())
            if d.get("tenant_id"):
                self._tenant_id = d["tenant_id"]
            self._access_token = d.get("access_token", "")
            self._expires_at = d.get("expires_at", 0.0)
            # Env var refresh token takes priority (Agent Hub credential)
            if not self._refresh_token:
                self._refresh_token = d.get("refresh_token", "")

    def _do_refresh(self):
        if not self._refresh_token or not self._tenant_id:
            return ""
        data = urllib.parse.urlencode({
            "client_id": CLIENT_ID,
            "grant_type": "refresh_token",
            "refresh_token": self._refresh_token,
            "scope": SCOPES,
        }).encode()
        req = urllib.request.Request(
            f"https://login.microsoftonline.com/{self._tenant_id}/oauth2/v2.0/token",
            data=data, method="POST"
        )
        try:
            with urllib.request.urlopen(req, timeout=30, context=self._ctx) as r:
                d = json.loads(r.read())
            self._access_token = d["access_token"]
            self._expires_at = time.time() + d.get("expires_in", 3600) - 60
            if "refresh_token" in d:
                self._refresh_token = d["refresh_token"]
            # Persist updated token to file if it exists (local runs)
            if TOKEN_FILE.exists():
                existing = json.loads(TOKEN_FILE.read_text())
                existing.update({
                    "access_token": self._access_token,
                    "expires_at": self._expires_at,
                    "refresh_token": self._refresh_token,
                })
                TOKEN_FILE.write_text(json.dumps(existing, indent=2))
            return self._access_token
        except Exception as e:
            print(f"ERROR: Token refresh failed: {e}", file=sys.stderr)
            return ""

    def token(self):
        if self._access_token and time.time() < self._expires_at:
            return self._access_token
        t = self._do_refresh()
        if not t:
            print("ERROR: No valid token. Run auth.py or set MS_REFRESH_TOKEN.", file=sys.stderr)
            sys.exit(1)
        return t

    def headers(self):
        return {
            "Authorization": f"Bearer {self.token()}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def get(self, path):
        url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
        req = urllib.request.Request(url, headers=self.headers())
        try:
            with urllib.request.urlopen(req, timeout=30, context=self._ctx) as r:
                return json.loads(r.read()), None
        except urllib.error.HTTPError as e:
            return None, f"HTTP {e.code}"
        except Exception as e:
            return None, str(e)

    def patch(self, path, body):
        url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
        req = urllib.request.Request(url, data=json.dumps(body).encode(),
                                     headers=self.headers(), method="PATCH")
        try:
            with urllib.request.urlopen(req, timeout=30, context=self._ctx) as r:
                return True, None
        except urllib.error.HTTPError as e:
            return False, f"HTTP {e.code}: {e.read().decode()[:200]}"
        except Exception as e:
            return False, str(e)


# ---------------------------------------------------------------------------
# Analytics helpers
# ---------------------------------------------------------------------------

def resolve_drive_ids(client, doc_name):
    """Search collateral site by filename → (drive_item_id, drive_id)."""
    q = urllib.parse.quote(doc_name)
    data, err = client.get(f"/sites/{COLLATERAL_SITE_ID}/drive/root/search(q='{q}')")
    if err or not data:
        return None, None
    for item in data.get("value", []):
        if item.get("name", "").lower() == doc_name.lower():
            return item["id"], item["parentReference"]["driveId"]
    if data.get("value"):
        item = data["value"][0]
        return item["id"], item["parentReference"]["driveId"]
    return None, None


def get_7d_analytics(client, drive_id, item_id):
    data, _ = client.get(f"/drives/{drive_id}/items/{item_id}/analytics/lastSevenDays")
    if not data:
        return 0, 0
    access = data.get("access", {})
    return access.get("actionCount", 0), access.get("actorCount", 0)


def categorize_from_path(folder_path):
    """Best-effort group from the file's folder path. Curated groups
    (Maintained / One-off / AAI Day '26) aren't folder-derivable, so those
    default to 'Uncategorized' for a human to tag."""
    p = folder_path or ""
    if "Training Modules/Workshop" in p:
        return "TME Workshop"
    if "Performance_Briefs" in p or "Performance Briefs" in p:
        return "Performance Briefs"
    if "HW PM Collateral" in p:
        return "HW PM Collateral"
    if "SW PM Collateral" in p:
        return "SW PM Collateral"
    if "Recordings" in p:
        return "Recordings"
    if "Product Management" in p or "Presentations" in p:
        return "Instinct SW Product"
    return "Uncategorized"


def get_item_meta(client, drive_id, item_id):
    """Fetch webUrl (evergreen file link), last-modified date, owner (createdBy),
    and a folder-derived category."""
    data, _ = client.get(
        f"/drives/{drive_id}/items/{item_id}"
        "?$select=webUrl,lastModifiedDateTime,createdBy,parentReference"
    )
    if not data:
        return None
    owner = data.get("createdBy", {}).get("user", {}).get("displayName", "")
    folder_path = data.get("parentReference", {}).get("path", "")
    return {
        "webUrl": data.get("webUrl", ""),
        "lastModified": data.get("lastModifiedDateTime", "")[:10],
        "owner": owner,
        "category": categorize_from_path(folder_path),
    }


def get_alltime(client, drive_id, item_id):
    """Lifetime analytics: returns (views, unique_users) over the file's whole
    history via the allTime endpoint."""
    data, _ = client.get(f"/drives/{drive_id}/items/{item_id}/analytics/allTime")
    if not data:
        return 0, 0
    access = data.get("access", {})
    return access.get("actionCount", 0), access.get("actorCount", 0)


def get_30d_views(client, drive_id, item_id):
    end = datetime.datetime.utcnow()
    start = end - datetime.timedelta(days=30)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    s = urllib.parse.quote(start.strftime(fmt))
    e = urllib.parse.quote(end.strftime(fmt))
    data, _ = client.get(
        f"/drives/{drive_id}/items/{item_id}/getActivitiesByInterval"
        f"(startDateTime='{s}',endDateTime='{e}',interval='day')"
    )
    if not data:
        return None
    return sum(day.get("access", {}).get("actionCount", 0) for day in data.get("value", []))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print(f"TME Content Analytics Refresh — {datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    client = TokenClient()
    # Trigger auth check early with a lightweight call
    me, err = client.get("/me?$select=displayName")
    if err:
        print(f"ERROR: Auth check failed: {err}")
        sys.exit(1)
    print(f"Authenticated as: {me.get('displayName', 'unknown')}\n")

    data, err = client.get(f"/sites/{LIST_SITE_ID}/lists/{LIST_ID}/items?$expand=fields&$top=100")
    if err:
        print(f"FATAL: Could not fetch list items: {err}")
        sys.exit(1)

    items = data.get("value", [])
    print(f"Found {len(items)} list item(s).\n")

    # Known column internal names — used to drop any field whose column
    # hasn't been created in the list UI yet (e.g. ViewsLifetime), so a
    # missing column can't fail the whole PATCH.
    cols_data, _ = client.get(f"/sites/{LIST_SITE_ID}/lists/{LIST_ID}/columns?$select=name")
    known_cols = {c.get("name") for c in (cols_data or {}).get("value", [])}

    stats = {"updated": 0, "resolved": 0, "no_analytics": 0, "errors": []}

    for item in items:
        item_id = item["id"]
        fields = item.get("fields", {})
        doc_name = fields.get("Title", "").strip()
        if not doc_name:
            continue

        drive_item_id = fields.get("DriveItemId", "").strip()
        drive_id = fields.get("DriveId", "").strip()

        # Auto-resolve missing IDs for newly added rows
        if not drive_item_id or not drive_id:
            print(f"  [{doc_name}] Resolving DriveItemId...")
            drive_item_id, drive_id = resolve_drive_ids(client, doc_name)
            if drive_item_id:
                ok, err = client.patch(
                    f"/sites/{LIST_SITE_ID}/lists/{LIST_ID}/items/{item_id}/fields",
                    {"DriveItemId": drive_item_id, "DriveId": drive_id}
                )
                if ok:
                    print(f"  [{doc_name}] IDs resolved and saved.")
                    stats["resolved"] += 1
                else:
                    print(f"  [{doc_name}] ERROR saving IDs: {err}")
                    stats["errors"].append(f"{doc_name}: ID save failed — {err}")
                    continue
            else:
                print(f"  [{doc_name}] Not found on collateral site — skipping.")
                stats["errors"].append(f"{doc_name}: not found on collateral site")
                continue

        views_7d, _ = get_7d_analytics(client, drive_id, drive_item_id)
        life_views, life_unique = get_alltime(client, drive_id, drive_item_id)

        update_fields = {
            "Views_x0028_last7d_x0029_": views_7d,
            "UniqueViews": life_unique,          # lifetime unique users
            "ViewsLifetime": life_views,         # lifetime total views
            "LastRefreshed": datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        }

        # Fill Location / LastUpdated / Owner / Category only when currently
        # empty (preserves hand-curated values on existing rows).
        if not fields.get("Location") or not fields.get("LastUpdated") \
                or not fields.get("Owner_x007c_Updatedby") \
                or not fields.get("Category"):
            meta = get_item_meta(client, drive_id, drive_item_id)
            if meta:
                if not fields.get("Location") and meta["webUrl"]:
                    update_fields["Location"] = meta["webUrl"]
                if not fields.get("LastUpdated") and meta["lastModified"]:
                    update_fields["LastUpdated"] = meta["lastModified"]
                if not fields.get("Owner_x007c_Updatedby") and meta["owner"]:
                    update_fields["Owner_x007c_Updatedby"] = meta["owner"]
                if not fields.get("Category") and meta["category"]:
                    update_fields["Category"] = meta["category"]

        # Fetch 30d views for every tracked row.
        v30 = get_30d_views(client, drive_id, drive_item_id)
        if v30 is not None:
            update_fields["Views_x0028_last30d_x0029_"] = v30

        if views_7d == 0 and life_views == 0:
            stats["no_analytics"] += 1

        # Drop any field whose column doesn't exist in the list yet.
        if known_cols:
            update_fields = {k: v for k, v in update_fields.items() if k in known_cols}

        ok, err = client.patch(
            f"/sites/{LIST_SITE_ID}/lists/{LIST_ID}/items/{item_id}/fields",
            update_fields
        )
        if ok:
            v30_str = f" 30d:{update_fields['Views_x0028_last30d_x0029_']}" if "Views_x0028_last30d_x0029_" in update_fields else ""
            life_str = f" life:{update_fields['ViewsLifetime']}" if "ViewsLifetime" in update_fields else ""
            print(f"  [{doc_name}] Updated — 7d:{views_7d} unique(lt):{life_unique}{v30_str}{life_str}")
            stats["updated"] += 1
        else:
            print(f"  [{doc_name}] ERROR: {err}")
            stats["errors"].append(f"{doc_name}: PATCH failed — {err}")

    print(f"\n{'='*60}")
    print(f"  Updated:         {stats['updated']}")
    print(f"  Newly resolved:  {stats['resolved']}")
    print(f"  Zero analytics:  {stats['no_analytics']}")
    if stats["errors"]:
        print(f"  Errors ({len(stats['errors'])}):")
        for e in stats["errors"]:
            print(f"    - {e}")
    else:
        print("  Errors:          none")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
