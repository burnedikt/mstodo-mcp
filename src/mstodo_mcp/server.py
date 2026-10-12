"""Microsoft To Do as a stdio MCP server.

The tools wrap the Graph calls in mstodo.py, so the sign-in, token cache and
list matching are the same as the `mstodo` CLI. Sign in once with
`mstodo login`, using the same MSTODO_CONFIG_DIR as the server.
"""

from types import SimpleNamespace

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from mstodo_mcp import mstodo

server = MCPServer(
    "microsoft-todo",
    instructions=(
        "Manages the user's Microsoft To Do. A project is a To Do list. "
        "Without a list name, tasks go to the default list. Convert relative "
        f"dates to YYYY-MM-DD in {mstodo.TIMEZONE}, and only set a due date or "
        "reminder when the user asked for one."
    ),
)


def run(fn, **args):
    try:
        return fn(SimpleNamespace(**args))
    except mstodo.Fail as e:
        if str(e).startswith(("not logged in", "token refresh failed")):
            raise ToolError(
                "Microsoft To Do is not signed in. Tell the user to run "
                "`mstodo login` once where this server runs; you cannot "
                "complete the sign-in yourself."
            ) from None
        raise ToolError(str(e)) from None


@server.tool()
def list_lists() -> list[dict]:
    """List all To Do lists (projects). `default` marks the default list."""
    return run(mstodo.cmd_lists)


@server.tool()
def list_tasks(list: str | None = None, include_completed: bool = False) -> list[dict]:
    """List open tasks, across all lists or in one list.

    `list` matches case-insensitively, then by prefix, then by substring.
    """
    return run(mstodo.cmd_tasks, list=list, all=include_completed)


@server.tool()
def add_task(
    title: str,
    list: str | None = None,
    create_list: bool = False,
    due: str | None = None,
    remind: str | None = None,
    note: str | None = None,
    important: bool = False,
) -> dict:
    """Add a task. Keep the title short and imperative; put detail in `note`.

    `list` is the project name; omit it for the default list. Set
    `create_list` only after the user confirmed that it is a new project.
    `due` is YYYY-MM-DD, `remind` is "YYYY-MM-DD HH:MM" local time.
    """
    return run(mstodo.cmd_add, title=title, list=list, create_list=create_list,
               due=due, remind=remind, note=note, important=important)


@server.tool()
def complete_task(task_id: str, list: str) -> dict:
    """Mark a task completed. Take `task_id` and `list` from list_tasks; never guess the id."""
    return run(mstodo.cmd_complete, task_id=task_id, list=list)


@server.tool()
def create_list(name: str) -> dict:
    """Create a To Do list (project). Only when the user asked for a new project."""
    return run(mstodo.cmd_create_list, name=name)


def main():
    server.run()
