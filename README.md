# office365-blade-mcp

Microsoft 365 Email, Calendar & Tasks via Microsoft Graph API. MCP server for the Sidereal platform.

## Features

- 39+ tools: email (read/write/delta), calendar (CRUD/freebusy), tasks (To Do + Planner)
- Token-efficient: `$select` field filtering, bodyPreview mode, pipe-delimited output
- Write-gated: destructive operations disabled by default
- Auth: device code (interactive) or client credentials (headless)
- Sidereal contracts: `email-v1`, `calendar-v1`, `tasks-v1`

## Quick Start

```bash
# Install
uv pip install -e .

# Configure
export O365_TENANT_ID="your-tenant-id"
export O365_CLIENT_ID="your-client-id"
export O365_AUTH_MODE="device_code"

# Run (stdio)
office365-blade-mcp

# Run (HTTP)
export O365_MCP_TRANSPORT=http
office365-blade-mcp
```

## License

MIT
