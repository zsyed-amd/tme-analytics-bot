#!/usr/bin/env python3
"""
TME Content Analytics — Monthly Refresh
Reads the TME Content Analytics SharePoint list, auto-resolves any rows
with missing DriveItemId, then updates View/Unique analytics from Graph API.

Required env vars:
  CLIENT_ID      — app registration client ID
  CLIENT_SECRET  — app registration client secret
  TENANT_ID      — AAD tenant ID

Optional env vars (fall back to hardcoded defaults):
  LIST_SITE_ID   — personal OneDrive site where the list lives
  LIST_ID        — GUID of the TME Content analytics list
  COLLATERAL_SITE_ID — site ID where the actual collateral files live
"""

import os, sys, json, datetime, urllib.parse
import urllib.request, urllib.error

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

TENANT_ID = os.getenv("TENANT_ID", "3dd8961f-e488-4e60-8e11-a82d994e183d")
CLIENT_ID = os.getenv("CLIENT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")

LIST_SITE_ID = os.getenv(
    "LIST_SITE_ID",
    "amdcloud-my.sharepoint.com,fc60e54f-ec1b-469b-93c2-d79471b2a67c,cd704538-2cf6-49ae-a1a4-4343ec2e3b38",
)
LIST_ID = os.getenv("LIST_ID", "977e2de1-9663-40ea-bbd8-d4fc80a63a11")
COLLATERAL_SITE_ID = os.getenv(
    "COLLATERAL_SITE_ID",
    "amdcloud.sharepoint.com,8a79450f-1df2-4861-8fe8-73ceec656271,ea2697f2-a242-4280-988a-09bdedf5919c",
)

GRAPH = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_token():
    if not CLIENT_ID or not CLIENT_SECRET:
        print("ERROR: CLIENT_ID and CLIENT_SECRET env vars are required for scheduled runs.")
        sys.exit(1)

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    body = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }).encode()
    req = urllib.request.Request(url, data=body, method="POST")
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read())["access_token"]


TOKEN = None

def headers():
    return {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def graph_get(path):
    url = path if path.startswith("http") else f"{GRAPH}{path}"
    req = urllib.request.Request(url, headers=headers())
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return json.loads(r.read()), None
    except urllib.error.HTTPError as e:
        return None, f"HTTP {e.code}"
    except Exception as e:
        return None, str(e)

def graph_patch(path, payload):
    url = path if path.startswith("http") else f"{GRAPH}{path}"
    data = json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data, method="PATCH", headers=headers())
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return True, None
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        return False, f"HTTP {e.code}: {body[:200]}"
    except Exception as e:
        return False, str(e)

# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def resolve_drive_ids(doc_name):
    """Search the collateral site for a file by name, return (drive_item_id, drive_id) or (None, None)."""
    q = urllib.parse.quote(doc_name)
    data, err = graph_get(f"/sites/{COLLATERAL_SITE_ID}/drive/root/search(q='{q}')")
    if err or not data:
        return None, None
    for item in data.get("value", []):
        if item.get("name", "").lower() == doc_name.lower():
            return item["id"], item["parentReference"]["driveId"]
    # Fuzzy fallback: first result if exact match not found
    if data.get("value"):
        item = data["value"][0]
        return item["id"], item["parentReference"]["driveId"]
    return None, None

def get_7d_analytics(drive_id, item_id):
    """Returns (views_7d, unique_views) or (0, 0) on failure."""
    data, err = graph_get(f"/drives/{drive_id}/items/{item_id}/analytics/lastSevenDays")
    if err or not data:
        return 0, 0
    access = data.get("access", {})
    return access.get("actionCount", 0), access.get("actorCount", 0)

def get_30d_views(drive_id, item_id):
    """Returns total view count over last 30 days, or None on failure."""
    end = datetime.datetime.utcnow()
    start = end - datetime.timedelta(days=30)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    start_str = urllib.parse.quote(start.strftime(fmt))
    end_str = urllib.parse.quote(end.strftime(fmt))
    path = f"/drives/{drive_id}/items/{item_id}/getActivitiesByInterval(startDateTime='{start_str}',endDateTime='{end_str}',interval='day')"
    data, err = graph_get(path)
    if err or not data:
        return None
    return sum(day.get("access", {}).get("actionCount", 0) for day in data.get("value", []))

def patch_list_item(item_id, fields):
    path = f"/sites/{LIST_SITE_ID}/lists/{LIST_ID}/items/{item_id}/fields"
    ok, err = graph_patch(path, fields)
    return ok, err

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    global TOKEN
    print(f"TME Content Analytics Refresh — {datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    print("Authenticating...")
    TOKEN = get_token()
    print("Token acquired.\n")

    # Fetch all list items
    data, err = graph_get(f"/sites/{LIST_SITE_ID}/lists/{LIST_ID}/items?$expand=fields&$top=100")
    if err:
        print(f"FATAL: Could not fetch list items: {err}")
        sys.exit(1)

    items = data.get("value", [])
    print(f"Found {len(items)} list item(s).\n")

    stats = {"updated": 0, "resolved": 0, "no_analytics": 0, "errors": []}

    for item in items:
        item_id = item["id"]
        fields = item.get("fields", {})
        doc_name = fields.get("Title", fields.get("DocumentName", "")).strip()

        if not doc_name:
            continue

        drive_item_id = fields.get("DriveItemId", "").strip()
        drive_id = fields.get("DriveId", "").strip()

        # Auto-resolve missing IDs
        if not drive_item_id or not drive_id:
            print(f"  [{doc_name}] DriveItemId missing — searching collateral site...")
            drive_item_id, drive_id = resolve_drive_ids(doc_name)
            if drive_item_id:
                ok, err = patch_list_item(item_id, {"DriveItemId": drive_item_id, "DriveId": drive_id})
                if ok:
                    print(f"  [{doc_name}] Resolved and saved IDs.")
                    stats["resolved"] += 1
                else:
                    print(f"  [{doc_name}] ERROR saving IDs: {err}")
                    stats["errors"].append(f"{doc_name}: ID save failed — {err}")
                    continue
            else:
                print(f"  [{doc_name}] Could not find file on collateral site — skipping.")
                stats["errors"].append(f"{doc_name}: not found on collateral site")
                continue

        # Pull analytics
        views_7d, unique_views = get_7d_analytics(drive_id, drive_item_id)
        has_30d = fields.get("Views30d") is not None  # only update 30d for rows that already had it

        update_fields = {
            "Views7d": views_7d,
            "UniqueViews": unique_views,
            "LastRefreshed": datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        }

        if has_30d:
            views_30d = get_30d_views(drive_id, drive_item_id)
            if views_30d is not None:
                update_fields["Views30d"] = views_30d

        if views_7d == 0 and unique_views == 0:
            stats["no_analytics"] += 1

        ok, err = patch_list_item(item_id, update_fields)
        if ok:
            print(f"  [{doc_name}] Updated — 7d:{views_7d} unique:{unique_views}" +
                  (f" 30d:{update_fields.get('Views30d','—')}" if has_30d else ""))
            stats["updated"] += 1
        else:
            print(f"  [{doc_name}] ERROR updating: {err}")
            stats["errors"].append(f"{doc_name}: PATCH failed — {err}")

    # Summary
    print(f"\n{'='*60}")
    print(f"  Updated:         {stats['updated']}")
    print(f"  Newly resolved:  {stats['resolved']}")
    print(f"  Zero analytics:  {stats['no_analytics']} (defaulted to 0)")
    if stats["errors"]:
        print(f"  Errors ({len(stats['errors'])}):")
        for e in stats["errors"]:
            print(f"    - {e}")
    else:
        print("  Errors:          none")
    print(f"{'='*60}")

if __name__ == "__main__":
    main()
