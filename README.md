# TME Content Analytics Bot

Monthly Agent Hub bot that refreshes view analytics for NDA collateral documents in the TME Content Analytics SharePoint list.

## Files

- `refresh_analytics.py` — the script the bot runs
- `AGENT.md` — system prompt for the Agent Hub Claude Code agent

## Setup

### 1. App registration (Azure Portal — one-time)

Create an Entra ID app registration named `TME-Content-Analytics-Bot` with these Graph API **Application** permissions:
- `Files.Read.All`
- `Sites.Read.All`
- `Sites.ReadWrite.All`

Request admin consent. Create a client secret and save it.

### 2. Agent Hub credentials (agenthub.amd.com → My Credentials)

Store these as credentials on the agent:
- `CLIENT_ID` — app registration client ID
- `CLIENT_SECRET` — app registration client secret

`TENANT_ID` defaults to `3dd8961f-e488-4e60-8e11-a82d994e183d` (hardcoded in script).

### 3. Create the Agent Hub agent

- Framework: **Claude Code**
- System prompt: contents of `AGENT.md`
- Repository: this repo URL
- Denied tools: `WebSearch`, `WebFetch`

### 4. Schedule

Cron: `0 8 1 * *` (8 AM UTC, 1st of every month)  
Autonomy mode: Fire & forget  
Permission mode: auto

## Adding new documents to track

1. Open the TME Content Analytics SharePoint list
2. Add a new row with `DocumentName`, `Location`, and `Owner` filled in
3. Leave `DriveItemId` and `DriveId` blank

The bot will automatically find the file on the NDA collateral site and start tracking it on the next monthly run.

## Running locally (for testing)

```bash
export CLIENT_ID=your-client-id
export CLIENT_SECRET=your-secret
python3 refresh_analytics.py
```
