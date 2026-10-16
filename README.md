# mstodo-mcp

[![M8ven Score](https://m8ven.ai/badge/mcp/burnedikt-mstodo-mcp-iv8sc2?v=bfd0739e7c7a870b1a888d0217d51b8b)](https://m8ven.ai/mcp/burnedikt-mstodo-mcp-iv8sc2?s=readme)

Microsoft To Do for a personal Microsoft account, through Microsoft Graph:

- `mstodo-mcp`: a stdio [MCP](https://modelcontextprotocol.io) server with the
  tools `list_lists`, `list_tasks`, `add_task`, `complete_task` and
  `create_list`
- `mstodo`: a CLI for the same operations, plus the one-time `login`

It signs in with the device-code flow and Microsoft's public "Microsoft Graph
Command Line Tools" client, so no Azure app registration is needed. Only
personal accounts (outlook.com, hotmail.com, live.com) are supported, not work
or school accounts. This is an unofficial project, not affiliated with
Microsoft.

## Quick start

Requires [uv](https://docs.astral.sh/uv/). Sign in once: open the URL that is
printed and enter the code.

```bash
uvx --from git+https://github.com/burnedikt/mstodo-mcp mstodo login
uvx --from git+https://github.com/burnedikt/mstodo-mcp mstodo lists
```

Then add the server to your MCP client. For Claude Code:

```bash
claude mcp add microsoft-todo -- uvx --from git+https://github.com/burnedikt/mstodo-mcp mstodo-mcp
```

For clients configured with JSON (Claude Desktop, Cursor, ...):

```json
{
  "mcpServers": {
    "microsoft-todo": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/burnedikt/mstodo-mcp", "mstodo-mcp"]
    }
  }
}
```

Append a tag such as `@v0.1.0` to the URL to pin a release. The server and
`mstodo login` must use the same config directory.

## CLI

Every command prints JSON.

```bash
mstodo lists                                  # all lists
mstodo tasks                                  # open tasks across all lists
mstodo tasks --list "Garden"                  # open tasks in one list
mstodo add "Buy screws" --list "Garden"       # add to a list
mstodo add "Call the tax advisor" --due 2026-10-09 --important
mstodo add "Pay rent" --remind "2026-11-01 08:00"
mstodo complete <task-id> --list "Garden"
mstodo create-list "Garden"
```

Without `--list`, tasks go to the default list. List names match
case-insensitively, then by prefix, then by substring.

## Configuration

| Variable            | Default                | Purpose                                      |
| ------------------- | ---------------------- | -------------------------------------------- |
| `MSTODO_CONFIG_DIR` | `~/.config/mstodo`     | Refresh token (`token.json`, mode 600)       |
| `MSTODO_CLIENT_ID`  | Graph CLI Tools client | Use your own public client registration      |
| `MSTODO_TIMEZONE`   | local time zone or UTC | IANA time zone for due dates and reminders   |

The sign-in requests `Tasks.ReadWrite` and `offline_access`. The refresh token
is renewed on use and expires after 90 days without use; run `mstodo login`
again then. Revoke access at <https://account.live.com/consent/Manage>.

`src/mstodo_mcp/mstodo.py` uses only the Python standard library, so it also
works as a single file, e.g. copied to `/usr/local/bin/mstodo`.

## License

MIT
