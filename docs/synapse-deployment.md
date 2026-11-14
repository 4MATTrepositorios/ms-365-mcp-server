# Synapse deployment (4MATT)

This fork runs as the **Microsoft 365 (Outlook)** connector of the Synapse portal.

## How auth works

- Synapse users already sign in with Microsoft (Supabase Auth, Azure provider, app `Supabase`,
  `<synapse-sso-app-id>`). Synapse keeps their delegated refresh token and mints a
  Graph access token per chat turn (`ms-token-refresh`).
- The Synapse omni-agent sends that **delegated Graph token** as `Authorization: Bearer` to `/mcp`.
  The server runs in plain HTTP mode (no `--obo`): the bearer goes straight to Microsoft Graph, so
  every call runs with the user's own permissions. No second sign-in, no token storage here.
- `MS365_MCP_ALLOWED_TENANTS` refuses (403) any bearer whose `tid` is not the 4MATT tenant.
  Graph still validates the token itself on every call.
- Scopes come from the tenant-wide admin consent on the `Supabase` app: Mail.ReadWrite, Mail.Send,
  Calendars.ReadWrite, Calendars.Read.Shared, Contacts.Read, MailboxSettings.Read (plus the
  portal's existing ones).

## Azure Container App

| Item  | Value                                                                                                                |
| ----- | -------------------------------------------------------------------------------------------------------------------- |
| App   | `ca-mcp-m365-synapse` (`<resource-group>`, env `cae-synapse`, brazilsouth)                                           |
| URL   | `https://<container-app-fqdn>/mcp`                                                                                   |
| Image | `ghcr.io/4mattrepositorios/ms-365-mcp-server:sha-<commit>` (private, pulled with the shared ghcr.io registry secret) |
| Args  | `--http 0.0.0.0:3000 --org-mode` (the tool list comes from `ENABLED_TOOLS`, not from a preset)                       |
| Scale | 0 to 3 replicas, 0.5 vCPU / 1 GiB                                                                                    |

Environment (no secrets):

```
MS365_MCP_ALLOWED_TENANTS=<4matt-tenant-id>
MS365_MCP_ALLOWED_SCOPES=<every Graph scope the enabled tools declare, space separated>
ENABLED_TOOLS=^(<tool>|<tool>|...)$   # exact names; must match the Synapse m365 tool_allowlist
MS365_MCP_CLIENT_ID=<synapse-sso-app-id>
MS365_MCP_TENANT_ID=<4matt-tenant-id>
MS365_MCP_DISABLE_DCR=true
MS365_MCP_TRUST_PROXY_HOPS=1
MS365_MCP_PUBLIC_URL=https://<container-app-fqdn>
MS365_MCP_MAX_TOP=25
LOG_LEVEL=info
```

## Release

```bash
docker build -t ghcr.io/4mattrepositorios/ms-365-mcp-server:sha-$(git rev-parse --short HEAD) .
docker push ghcr.io/4mattrepositorios/ms-365-mcp-server:sha-$(git rev-parse --short HEAD)
az containerapp update -n ca-mcp-m365-synapse -g <resource-group> \
  --image ghcr.io/4mattrepositorios/ms-365-mcp-server:sha-$(git rev-parse --short HEAD)
```

`az containerapp create/update --args` cannot pass values that start with `--`; change args with
`az containerapp update --yaml` (template only; registry secrets are kept).

## Smoke test

- `GET /` returns 200.
- `POST /mcp` without a bearer returns 401; with a token from another tenant, 403.

The Synapse side (connector row `m365`, `microsoft_user` auth type, pilot role `outlook`) is
documented in the Synapse repo under DB-195, OMNI-150 and CMP-92.

## Tool rollout

The tool list lives in two places that must match: `ENABLED_TOOLS` on the container (what the
server exposes) and the `tool_allowlist` of the `m365` row in Synapse (what the chat offers).
On 14/11/2026 it went from the 28 Outlook tools to 93 (Outlook, People, Teams chats, channels and
meetings, OneNote, Microsoft 365 search). New Graph scopes are added to the SSO app's admin
consent before a block is enabled; a tool whose scope is not consented answers 403.
