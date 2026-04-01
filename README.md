# office365-blade-mcp

A token-efficient Microsoft 365 MCP server for email, calendar, and tasks -- all through one connection.

Built on the [Model Context Protocol](https://modelcontextprotocol.io/) and Microsoft Graph API. Ships 38 tools with pipe-delimited output, field selection, null omission, write-gating, and dual auth. Designed to give LLMs maximum signal per token.

---

## Why This Over Microsoft's Official Graph MCP Servers?

| | **office365-blade-mcp** | Microsoft reference servers |
|---|---|---|
| **Scope** | Email + Calendar + Tasks in ONE server | Split across separate repos |
| **Output** | Pipe-delimited, field-selected, nulls omitted | Raw JSON blobs (10-50x more tokens) |
| **Auth** | Dual: device code (interactive) + client credentials (headless/CI) | Single flow per server |
| **Safety** | Write-gated by default, delete confirmation, credential scrubbing | Varies |
| **Sync** | Delta tokens for incremental email sync | Full re-fetch |
| **Convenience** | `cal_today`, `cal_week`, `task_today`, `task_inbox` | Build your own queries |
| **Batch** | `email_bulk` (up to 50 ops), `cal_batch` (multi-calendar) | Individual calls |

The output format alone justifies the switch. A 20-email search result that costs ~8K tokens from raw Graph JSON costs ~800 tokens here.

---

## Features

- **38 tools** across email, calendar, To Do, and Planner
- **Token-efficient output** -- pipe-delimited tables with `$select` optimization and null omission
- **Write-gated by default** -- read-only until you explicitly enable writes
- **Dual authentication** -- device code flow for interactive use, client credentials for daemons and CI
- **Incremental sync** -- delta tokens for email change tracking without re-fetching
- **Batch operations** -- bulk email actions (up to 50), multi-calendar event fetch
- **Purpose-built views** -- `cal_today`, `cal_week`, `task_today`, `task_inbox` for common queries
- **FastMCP 2.0** -- stdio and HTTP transports

---

## Quick Start

```bash
# Install
uvx office365-blade-mcp

# Or clone and install
git clone https://github.com/groupthink-dev/office365-blade-mcp.git
cd office365-blade-mcp
uv sync

# Configure
export O365_TENANT_ID="your-tenant-id"
export O365_CLIENT_ID="your-client-id"
export O365_AUTH_MODE="device_code"

# Run (stdio)
uv run office365-blade-mcp
```

On first run with `device_code` auth, you'll be prompted to authenticate in your browser.

---

## Tools

### Email -- Meta

| Tool | Description |
|------|-------------|
| `o365_info` | Account and tenant info, health check |
| `email_folders` | List mail folders with unread/total counts |

### Email -- Read

| Tool | Description |
|------|-------------|
| `email_search` | Search with OData filters and text queries |
| `email_read` | Full email content with body |
| `email_thread` | Conversation view -- all messages in a thread |
| `email_snippets` | Search with context excerpts for quick triage |

### Email -- Attachments

| Tool | Description |
|------|-------------|
| `email_attachments` | List attachments for an email |

### Email -- Sync

| Tool | Description |
|------|-------------|
| `email_state` | Get a delta token for current mailbox state |
| `email_changes` | Incremental changes since a previous delta token |

### Email -- Write (gated)

| Tool | Description |
|------|-------------|
| `email_send` | Send a new email |
| `email_reply` | Reply to an email |
| `email_forward` | Forward an email |
| `email_flag` | Flag/unflag an email |
| `email_move` | Move email to a folder |
| `email_delete` | Delete an email (requires confirmation) |
| `email_bulk` | Batch operations on up to 50 emails |

### Calendar -- Read

| Tool | Description |
|------|-------------|
| `cal_calendars` | List available calendars |
| `cal_events` | Events in a date range |
| `cal_event` | Full event detail by ID |
| `cal_search` | Search events by text |
| `cal_today` | Today's events across all calendars |
| `cal_week` | This week's events across all calendars |

### Calendar -- Write (gated)

| Tool | Description |
|------|-------------|
| `cal_freebusy` | Check availability for scheduling |
| `cal_batch` | Fetch events from multiple calendars in one call |
| `cal_respond` | Accept, decline, or tentatively accept an invite |
| `cal_create` | Create a new calendar event |
| `cal_update` | Update an existing event |
| `cal_delete` | Delete an event (requires confirmation) |

### Tasks -- To Do

| Tool | Description |
|------|-------------|
| `task_lists` | List all task lists |
| `task_list` | Tasks in a specific list |
| `task_search` | Search tasks by text |
| `task_today` | Tasks due today |
| `task_inbox` | Tasks in the default list |
| `task_create` | Create a task (gated) |
| `task_update` | Update a task (gated) |
| `task_complete` | Mark a task complete (gated) |

### Tasks -- Planner

| Tool | Description |
|------|-------------|
| `planner_plans` | List Planner plans |
| `planner_tasks` | Tasks in a Planner plan |

---

## Output Format

All read tools return pipe-delimited, field-selected output with null fields omitted. This keeps LLM context tight.

**Raw Graph API response** (~400 tokens):
```json
{"id": "AAMk...", "subject": "Q3 Review", "from": {"emailAddress": {"name": "Alice", "address": "alice@contoso.com"}}, "receivedDateTime": "2026-03-28T14:30:00Z", "isRead": true, "importance": "normal", "hasAttachments": false, "bodyPreview": "Hi team, please review the attached...", "flag": {"flagStatus": "notFlagged"}, "categories": [], ...}
```

**office365-blade-mcp response** (~40 tokens):
```
AAMk...|Q3 Review|Alice <alice@contoso.com>|2026-03-28 14:30|read|Hi team, please review the attached...
```

---

## Authentication

### Device Code (interactive)

For desktop use and development. Prompts for browser-based login.

```bash
export O365_AUTH_MODE="device_code"
export O365_TENANT_ID="your-tenant-id"
export O365_CLIENT_ID="your-client-id"
```

### Client Credentials (headless)

For daemons, CI pipelines, and unattended operation. Requires a client secret and admin-consented app registration.

```bash
export O365_AUTH_MODE="client_credentials"
export O365_TENANT_ID="your-tenant-id"
export O365_CLIENT_ID="your-client-id"
export O365_CLIENT_SECRET="your-client-secret"
```

---

## Security Model

| Layer | Behaviour |
|-------|-----------|
| **Write gate** | All write/delete tools disabled by default. Enable with `O365_WRITE_ENABLED=true` |
| **Delete confirmation** | Delete operations require explicit confirmation parameters |
| **Credential scrubbing** | Bearer tokens and secrets never appear in tool output or logs |
| **Bearer auth** | Optional API token for HTTP transport via `O365_MCP_API_TOKEN` |
| **Minimal scopes** | Requests use `$select` to fetch only the fields needed |

---

## Claude Desktop Config

```json
{
  "mcpServers": {
    "office365": {
      "command": "uvx",
      "args": ["office365-blade-mcp"],
      "env": {
        "O365_TENANT_ID": "your-tenant-id",
        "O365_CLIENT_ID": "your-client-id",
        "O365_AUTH_MODE": "device_code"
      }
    }
  }
}
```

## Claude Code Config

```json
{
  "mcpServers": {
    "office365": {
      "command": "uvx",
      "args": ["office365-blade-mcp"],
      "env": {
        "O365_TENANT_ID": "your-tenant-id",
        "O365_CLIENT_ID": "your-client-id",
        "O365_AUTH_MODE": "device_code",
        "O365_WRITE_ENABLED": "true"
      }
    }
  }
}
```

---

## Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `O365_TENANT_ID` | Yes | -- | Azure AD tenant ID |
| `O365_CLIENT_ID` | Yes | -- | App registration client ID |
| `O365_CLIENT_SECRET` | For client_credentials | -- | App registration client secret |
| `O365_AUTH_MODE` | No | `device_code` | `device_code` or `client_credentials` |
| `O365_WRITE_ENABLED` | No | `false` | Enable write/delete tools |
| `O365_MCP_TRANSPORT` | No | `stdio` | `stdio` or `http` |
| `O365_MCP_HOST` | No | `127.0.0.1` | HTTP transport bind address |
| `O365_MCP_PORT` | No | `8000` | HTTP transport port |
| `O365_MCP_API_TOKEN` | No | -- | Bearer token for HTTP transport auth |

---

## Architecture

```
src/office365_blade_mcp/
├── server.py       # FastMCP 2.0 server, 38 @mcp.tool decorators
├── client.py       # GraphClient: dual auth, token refresh, $select optimization
├── formatters.py   # Token-efficient output: pipe-delimited, null omission, field selection
├── models.py       # Config dataclass, write-gate logic, constants
└── auth.py         # Device code + client credentials flows, bearer middleware
```

---

## Development

```bash
# Install dependencies
uv sync

# Run locally (stdio)
uv run office365-blade-mcp

# Run with HTTP transport
O365_MCP_TRANSPORT=http uv run office365-blade-mcp

# Lint
uv run ruff check .

# Format
uv run ruff format .
```

---

## License

MIT
