# Outlook MCP Server

Public **Outlook / Microsoft 365 mail** MCP server — Streamable HTTP + OAuth 2.1,
same architecture as the Salesforce MCP Server. One deployed URL works with:

- **Archon AI portal** (Managed MCP via Anthropic / OpenAI)
- **Claude.ai** custom connectors
- **ChatGPT** custom MCP
- Any standalone MCP client (n8n, Cursor, custom apps) via the PATH B flow

## Tools (11)

| Tool | Kind | Description |
|---|---|---|
| `getProfile` | read | Connected mailbox owner (name, email) |
| `listEmails` | read | List a folder's messages, filters: unread, sender |
| `readEmail` | read | Full body of one message |
| `searchEmails` | read | Full-text search across all folders (KQL) |
| `listFolders` | read | Folder tree with unread counts |
| `getAttachments` | read | Attachment names/types/sizes |
| `sendEmail` | write | Send a new email |
| `replyEmail` | write | Reply / reply-all in-thread |
| `createDraft` | write | Draft without sending |
| `moveEmail` | write | Move to folder / archive / trash |
| `markEmail` | write | Read/unread + follow-up flag |

## Setup

### 1. Azure app registration (one-time)

1. https://entra.microsoft.com → **App registrations** → **New registration**
2. Supported account types: *Accounts in any organizational directory and
   personal Microsoft accounts*
3. **Redirect URIs** (type Web):
   - `https://claude.ai/api/mcp/auth_callback` (Claude connector)
   - `<BASE_URL>/auth/callback` (PATH B standalone)
4. **API permissions** → Microsoft Graph → *Delegated*:
   `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `User.Read`, `offline_access`
5. **Certificates & secrets** → New client secret → copy the **Value**

### 2. Run

```bash
cp .env.example .env      # fill in MS_CLIENT_ID / MS_CLIENT_SECRET / BASE_URL
npm install
npm run dev               # http://localhost:3100
```

### 3. Deploy (Render)

- Build command: `npm install`
- Start command: `npm start`
- Env vars: `BASE_URL` (the Render URL), `MS_CLIENT_ID`, `MS_CLIENT_SECRET`,
  `MS_TENANT=common`, `MS_REDIRECT_URI=<BASE_URL>/auth/callback`
- After deploy, add `<BASE_URL>/auth/callback` to the Azure app's redirect URIs.

## Auth paths

**PATH A — Bearer header (portals, Claude connector).** The client sends
`Authorization: Bearer <graph access token>` on `/mcp`. Tokens acquired
through this server's `/token` proxy get vaulted, enabling server-side
refresh when Graph returns 401 (Graph tokens live ~1 hour).

**PATH B — session flow (standalone clients).** Visit
`<BASE_URL>/auth?session=<id>` → Microsoft login → tokens stored server-side
keyed by the stable Microsoft user id.

The `/authorize` and `/token` proxies force the real Azure client credentials
so Dynamic Client Registration (Claude invents a random client id) still
works — the only requirement is the client's redirect URI being whitelisted
on the Azure app.
