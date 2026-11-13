# mstodo-mcp

[![M8ven Score](https://m8ven.ai/badge/mcp/burnedikt-mstodo-mcp-iv8sc2?v=bfd0739e7c7a870b1a888d0217d51b8b)](https://m8ven.ai/mcp/burnedikt-mstodo-mcp-iv8sc2?s=readme)

Microsoft To Do for a personal Microsoft account, through Microsoft Graph:

- `mstodo-mcp`: a stdio [MCP](https://modelcontextprotocol.io) server with the
  tools `list_lists`, `list_tasks`, `list_my_day`, `add_task`,
  `complete_task` and `create_list`
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

Append a tag such as `@v0.2.0` to the URL to pin a release. The server and
`mstodo login` must use the same config directory.

## CLI

Every command prints JSON.

```bash
mstodo lists                                  # all lists
mstodo tasks                                  # open tasks across all lists
mstodo tasks --list "Garden"                  # open tasks in one list
mstodo my-day                                 # tasks in My Day today
mstodo add "Buy screws" --list "Garden"       # add to a list
mstodo add "Call the tax advisor" --due 2026-10-09 --important
mstodo add "Pay rent" --remind "2026-11-01 08:00"
mstodo complete <task-id> --list "Garden"
mstodo create-list "Garden"
```

Without `--list`, tasks go to the default list. List names match
case-insensitively, then by prefix, then by substring.

### My Day

Microsoft Graph does not expose My Day, so `mstodo my-day` and `list_my_day`
read it from the undocumented API that the To Do apps themselves use
(`substrate.office.com/todob2`). It works today, but Microsoft may change or
remove it without notice; the other commands only use Graph and are not
affected. "Today" is the current date in `MSTODO_TIMEZONE`. Add `--all` to
include tasks already completed today.

If you signed in before My Day support existed, run `mstodo login` again once
to grant access to that API.

## Configuration

| Variable            | Default                | Purpose                                      |
| ------------------- | ---------------------- | -------------------------------------------- |
| `MSTODO_CONFIG_DIR` | `~/.config/mstodo`     | Refresh token (`token.json`, mode 600)       |
| `MSTODO_CLIENT_ID`  | Graph CLI Tools client | Your own app registration, see below         |
| `MSTODO_TIMEZONE`   | local time zone or UTC | IANA time zone for due dates and reminders   |

The sign-in requests `Tasks.ReadWrite` for Microsoft Graph and for Outlook
(the To Do apps' API, used for My Day), plus `offline_access`. The refresh token
is renewed on use and expires after 90 days without use; run `mstodo login`
again then. Revoke access at <https://account.live.com/consent/Manage>.

### Your own app registration

By default the sign-in borrows Microsoft's "Microsoft Graph Command Line Tools"
client. Microsoft retired the Graph CLI itself in August 2026; the client still
works, but if it ever stops, or you prefer a client you control, register your
own:

1. In the [Entra admin center](https://entra.microsoft.com), go to **App
   registrations > New registration**. Choose **Personal Microsoft accounts
   only** and leave the redirect URI empty.
2. Under **Authentication**, set **Allow public client flows** to **Yes**.
3. Under **API permissions**, add the delegated Microsoft Graph permissions
   `Tasks.ReadWrite` and `offline_access`.
4. Sign in with the **Application (client) ID** from the overview page:

```bash
mstodo login --client-id <application-client-id>
```

The client ID is saved with the token, so the MCP server and later commands
need no further setting. Setting `MSTODO_CLIENT_ID`, or writing the ID to
`~/.config/mstodo/client_id`, instead makes it the default for `mstodo login`.
To go back to Microsoft's client, unset those and run `mstodo login` again.

My Day needs a second permission, for the To Do apps' own API, that the portal
does not offer. If the sign-in with your app fails because of it, run
`mstodo login --client-id <id> --no-my-day`: everything except My Day then
works.

`src/mstodo_mcp/mstodo.py` uses only the Python standard library, so it also
works as a single file, e.g. copied to `/usr/local/bin/mstodo`.

## License

MIT
