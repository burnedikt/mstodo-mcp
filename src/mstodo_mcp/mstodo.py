#!/usr/bin/env python3
"""Minimal Microsoft To Do CLI.

Talks to Microsoft Graph with a delegated token for a personal Microsoft
account. Standard library only, so this file also runs on its own, e.g.
copied to /usr/local/bin/mstodo. The MCP server (server.py) wraps the same
functions.

    mstodo login                         # one-time device-code sign-in
    mstodo lists
    mstodo tasks [--list NAME] [--all]
    mstodo add "Title" [--list NAME] [--create-list] [--due YYYY-MM-DD]
                       [--remind "YYYY-MM-DD HH:MM"] [--note TEXT] [--important]
    mstodo complete TASK_ID --list NAME
    mstodo create-list NAME

Output is JSON on stdout; errors are JSON on stderr with a non-zero exit.
Config: MSTODO_CLIENT_ID or ~/.config/mstodo/client_id (public client for
personal accounts with delegated Tasks.ReadWrite; defaults to Microsoft's
"Microsoft Graph Command Line Tools"). The refresh token is cached in
~/.config/mstodo/token.json (mode 600). MSTODO_CONFIG_DIR moves both.
Dates and reminders use MSTODO_TIMEZONE, else the local IANA time zone.
"""

import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = "Tasks.ReadWrite offline_access"


def local_timezone():
    """IANA name of the local time zone (Graph accepts these), else UTC."""
    name = os.environ.get("TZ", "").lstrip(":")
    if not name:
        _, _, name = os.path.realpath("/etc/localtime").partition("/zoneinfo/")
    return name or "UTC"


TIMEZONE = os.environ.get("MSTODO_TIMEZONE") or local_timezone()
CONFIG_DIR = Path(os.environ.get("MSTODO_CONFIG_DIR", Path.home() / ".config" / "mstodo"))
TOKEN_FILE = CONFIG_DIR / "token.json"
# Microsoft Graph Command Line Tools: a public client that accepts personal
# accounts and the device-code flow, so no app registration is needed.
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"


class Fail(Exception):
    pass


def client_id():
    cid = os.environ.get("MSTODO_CLIENT_ID")
    if not cid:
        cfg = CONFIG_DIR / "client_id"
        if cfg.exists():
            cid = cfg.read_text().strip()
    return cid or DEFAULT_CLIENT_ID


def post_form(url, fields):
    data = urllib.parse.urlencode(fields).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.load(resp)
    except urllib.error.HTTPError as e:
        return json.load(e)


def save_token(tok):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    tok["expires_at"] = time.time() + int(tok.get("expires_in", 3600)) - 60
    tmp = TOKEN_FILE.with_suffix(".tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump(tok, f)
    os.replace(tmp, TOKEN_FILE)


def login():
    flow = post_form(f"{AUTHORITY}/devicecode", {"client_id": client_id(), "scope": SCOPES})
    if "device_code" not in flow:
        raise Fail(f"device code request failed: {flow.get('error_description', flow)}")
    print(flow["message"], file=sys.stderr, flush=True)
    deadline = time.time() + int(flow.get("expires_in", 900))
    interval = int(flow.get("interval", 5))
    while time.time() < deadline:
        time.sleep(interval)
        tok = post_form(f"{AUTHORITY}/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id(),
            "device_code": flow["device_code"],
        })
        if "access_token" in tok:
            save_token(tok)
            return {"status": "logged_in"}
        if tok.get("error") == "slow_down":
            interval += 5
        elif tok.get("error") != "authorization_pending":
            raise Fail(f"login failed: {tok.get('error_description', tok)}")
    raise Fail("login timed out")


def access_token():
    if not TOKEN_FILE.exists():
        raise Fail("not logged in: run `mstodo login` once (needs the user to enter a code)")
    tok = json.loads(TOKEN_FILE.read_text())
    if time.time() < tok.get("expires_at", 0):
        return tok["access_token"]
    new = post_form(f"{AUTHORITY}/token", {
        "grant_type": "refresh_token",
        "client_id": client_id(),
        "refresh_token": tok["refresh_token"],
        "scope": SCOPES,
    })
    if "access_token" not in new:
        raise Fail(f"token refresh failed, run `mstodo login` again: {new.get('error_description', new)}")
    new.setdefault("refresh_token", tok["refresh_token"])
    save_token(new)
    return new["access_token"]


def graph(method, path, body=None):
    url = path if path.startswith("http") else f"{GRAPH}{path}"
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {access_token()}")
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read()
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        detail = e.read().decode(errors="replace")
        raise Fail(f"Graph {method} {path} -> {e.code}: {detail[:500]}")


def paged(path):
    items, url = [], path
    while url:
        page = graph("GET", url)
        items.extend(page.get("value", []))
        url = page.get("@odata.nextLink")
    return items


def all_lists():
    return paged("/me/todo/lists")


def find_list(name, create=False):
    lists = all_lists()
    if not name:
        default = [l for l in lists if l.get("wellknownListName") == "defaultList"]
        if default:
            return default[0]
        raise Fail("no default list found")
    wanted = name.casefold().strip()
    for match in (
        lambda n: n == wanted,
        lambda n: n.startswith(wanted),
        lambda n: wanted in n,
    ):
        hits = [l for l in lists if match(l["displayName"].casefold())]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            raise Fail(f"list name '{name}' is ambiguous: {[l['displayName'] for l in hits]}")
    if create:
        return graph("POST", "/me/todo/lists", {"displayName": name})
    raise Fail(f"no list named '{name}'. Existing lists: {[l['displayName'] for l in lists]}. "
               "Pass --create-list to create it.")


def slim_task(t, list_name):
    out = {"id": t["id"], "title": t["title"], "list": list_name, "status": t["status"]}
    if t.get("importance") == "high":
        out["important"] = True
    if t.get("dueDateTime"):
        out["due"] = t["dueDateTime"]["dateTime"][:10]
    if t.get("body", {}).get("content", "").strip():
        out["note"] = t["body"]["content"].strip()[:300]
    return out


def cmd_lists(_):
    return [{"id": l["id"], "name": l["displayName"], "default": l.get("wellknownListName") == "defaultList"}
            for l in all_lists()]


def cmd_tasks(a):
    targets = [find_list(a.list)] if a.list else all_lists()
    query = "" if a.all else "?$filter=" + urllib.parse.quote("status ne 'completed'")
    out = []
    for l in targets:
        for t in paged(f"/me/todo/lists/{l['id']}/tasks{query}"):
            out.append(slim_task(t, l["displayName"]))
    return out


def cmd_add(a):
    l = find_list(a.list, create=a.create_list)
    task = {"title": a.title}
    if a.note:
        task["body"] = {"content": a.note, "contentType": "text"}
    if a.important:
        task["importance"] = "high"
    if a.due:
        task["dueDateTime"] = {"dateTime": f"{a.due}T00:00:00", "timeZone": TIMEZONE}
    if a.remind:
        date, _, clock = a.remind.partition(" ")
        task["isReminderOn"] = True
        task["reminderDateTime"] = {"dateTime": f"{date}T{clock or '09:00'}:00", "timeZone": TIMEZONE}
    created = graph("POST", f"/me/todo/lists/{l['id']}/tasks", task)
    return slim_task(created, l["displayName"])


def cmd_complete(a):
    l = find_list(a.list)
    done = graph("PATCH", f"/me/todo/lists/{l['id']}/tasks/{a.task_id}", {"status": "completed"})
    return slim_task(done, l["displayName"])


def cmd_create_list(a):
    created = graph("POST", "/me/todo/lists", {"displayName": a.name})
    return {"id": created["id"], "name": created["displayName"]}


def main():
    p = argparse.ArgumentParser(prog="mstodo", description="Microsoft To Do via Microsoft Graph")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("login").set_defaults(fn=lambda _: login())
    sub.add_parser("lists").set_defaults(fn=cmd_lists)
    t = sub.add_parser("tasks")
    t.add_argument("--list")
    t.add_argument("--all", action="store_true", help="include completed tasks")
    t.set_defaults(fn=cmd_tasks)
    ad = sub.add_parser("add")
    ad.add_argument("title")
    ad.add_argument("--list", help="list (project) name; default list if omitted")
    ad.add_argument("--create-list", action="store_true")
    ad.add_argument("--due", help="YYYY-MM-DD")
    ad.add_argument("--remind", help="'YYYY-MM-DD HH:MM' local time")
    ad.add_argument("--note")
    ad.add_argument("--important", action="store_true")
    ad.set_defaults(fn=cmd_add)
    c = sub.add_parser("complete")
    c.add_argument("task_id")
    c.add_argument("--list", required=True)
    c.set_defaults(fn=cmd_complete)
    cl = sub.add_parser("create-list")
    cl.add_argument("name")
    cl.set_defaults(fn=cmd_create_list)
    a = p.parse_args()
    try:
        print(json.dumps(a.fn(a), ensure_ascii=False))
    except Fail as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
