# TME Content Analytics Bot

Monthly Agent Hub bot that refreshes view analytics for NDA collateral documents in the TME Content Analytics SharePoint list.

**No app registration or admin consent required.** Uses the same pre-authorized Teams Desktop client as the m365-* skills.

## Files

- `refresh_analytics.py` — the script the bot runs
- `AGENT.md` — system prompt for the Agent Hub Claude Code agent

## One-time Setup

### 1. Authenticate locally (one-time)

```bash
python3 /home/zsyed/.claude/skills/m365-teams/scripts/auth.py
# Follow the device code flow — sign in with your AMD Okta account
```

This saves a token to `~/.config/microsoft-graph/token.json`.

### 2. Extract the refresh token for Agent Hub

```bash
python3 -c "import json; d=json.load(open('/root/.config/microsoft-graph/token.json')); print(d['refresh_token'])"
```

Copy the output — this is your `MS_REFRESH_TOKEN`.

### 3. Create the Agent Hub agent (agenthub.amd.com → Create Agent)

| Field | Value |
|---|---|
| Name | `TME Content Analytics Refresh` |
| Framework | Claude Code |
| System prompt | Upload `AGENT.md` |
| Repository | `https://github.com/zsyed-amd/tme-analytics-bot` |
| Denied tools | WebSearch, WebFetch |

### 4. Store credential in Agent Hub (My Credentials)

| Key | Value |
|---|---|
| `MS_REFRESH_TOKEN` | *(the refresh token from step 2)* |

### 5. Schedule

- Cron: `0 8 1 * *` (8 AM UTC, 1st of every month)
- Autonomy mode: Fire & forget
- Permission mode: auto

## Adding new documents to track

1. Add a row to the SharePoint List with `Title`, `Location`, and `Owner | Updated by`
2. Leave `DriveItemId` and `DriveId` blank
3. Bot resolves the IDs and starts tracking on next monthly run

## Running locally

```bash
# Uses token from ~/.config/microsoft-graph/token.json automatically
python3 refresh_analytics.py
```

## Token maintenance

Refresh tokens stay alive as long as the bot runs at least once every 90 days (monthly runs keep it perpetually fresh). If the token ever expires, re-run `auth.py` locally, extract the new refresh token, and update the Agent Hub credential.
