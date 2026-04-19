#!/usr/bin/env python3
"""
Google Calendar Webhook Manager CLI
Manages push notification channels (webhooks) for Google Calendar events.

Setup:
  pip install google-auth google-auth-oauthlib google-api-python-client

Usage:
    # List all tracked webhooks (with expiry status)
    python gcal_webhooks.py list

    # Create a webhook for the primary calendar
    python gcal_webhooks.py create --webhook-url https://your.server/webhook
    python gcal_webhooks.py create --calendar-id <id> --webhook-url https://your.server/webhook
    python gcal_webhooks.py create --calendar-id <id> --webhook-url https://your.server/webhook --token mysecret

    # Refresh a single expiring webhook (stop old → create new)
    python gcal_webhooks.py refresh --channel-id <id> --resource-id <rid> --calendar-id <id> --webhook-url https://your.server/webhook

    # Refresh ALL expired webhooks automatically
    python gcal_webhooks.py refresh-all
    python gcal_webhooks.py refresh-all --all   # force refresh even non-expired

    # Delete a webhook
    python gcal_webhooks.py delete --channel-id <id>
    python gcal_webhooks.py delete --channel-id <id> --resource-id <id>
"""

import argparse
import json
import os
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

# ── Auth / API ────────────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/calendar"]
TOKEN_FILE = Path("token.json")
CREDENTIALS_FILE = Path("credentials.json")
STATE_FILE = Path("webhook_state.json")   # local record of active channels

def _run_local_server_interruptible(flow):
    """
    Spin up a temporary localhost HTTP server to receive the OAuth callback.
    Ctrl+C cancels cleanly. Tries ports 8080, 8090, 8888, 9000 in order.
    Add ALL of these to Authorized redirect URIs in Google Cloud Console:
      http://localhost:8080
      http://localhost:8090
      http://localhost:8888
      http://localhost:9000
    """
    import webbrowser
    import threading
    import socketserver
    from http.server import BaseHTTPRequestHandler
    from urllib.parse import urlparse, parse_qs

    PORTS = [8080, 8090, 8888, 9000]
    auth_code = []          # mutable container so the handler can write to it
    shutdown_event = threading.Event()

    class CallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            params = parse_qs(urlparse(self.path).query)
            if "code" in params:
                auth_code.append(params["code"][0])
                body = b"<h2>Authentication successful! You can close this tab.</h2>"
            elif "error" in params:
                auth_code.append(None)
                err = params["error"][0]
                body = f"<h2>Authentication failed: {err}</h2>".encode()
            else:
                body = b"<h2>Waiting...</h2>"
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(body)
            shutdown_event.set()

        def log_message(self, *args):
            pass  # suppress access log noise

    # Find a free port
    server = None
    port = None
    for p in PORTS:
        try:
            server = socketserver.TCPServer(("localhost", p), CallbackHandler)
            server.allow_reuse_address = True
            port = p
            break
        except OSError:
            continue

    if server is None:
        print(f"ERROR: Could not bind to any of {PORTS}. Free one of those ports and retry.")
        sys.exit(1)

    redirect_uri = f"http://localhost:{port}"
    flow.redirect_uri = redirect_uri
    auth_url, _ = flow.authorization_url(prompt="consent")

    print(f"Starting local OAuth callback server on {redirect_uri}")
    print(f"Make sure  {redirect_uri}  is in your Authorized redirect URIs in Google Cloud Console.")
    print("Opening browser for Google sign-in...")
    print(f"  {auth_url}")
    webbrowser.open(auth_url)
    print("Waiting for approval... (Ctrl+C to cancel)")

    # Run server in a background thread so Ctrl+C reaches the main thread
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()

    try:
        while not shutdown_event.is_set():
            shutdown_event.wait(timeout=0.5)
    except KeyboardInterrupt:
        print("Authentication cancelled.")
        server.shutdown()
        sys.exit(0)
    finally:
        server.shutdown()

    if not auth_code or auth_code[0] is None:
        print("Authentication failed or was denied.")
        sys.exit(1)

    flow.fetch_token(code=auth_code[0])
    return flow.credentials


def get_service():
    """Authenticate and return a Google Calendar API service object."""
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError:
        print("ERROR: Required packages not installed.")
        print("Run: pip install google-auth google-auth-oauthlib google-api-python-client")
        sys.exit(1)

    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request as GoogleRequest
            creds.refresh(GoogleRequest())
        else:
            if not CREDENTIALS_FILE.exists():
                print(f"ERROR: {CREDENTIALS_FILE} not found.")
                print("Download it from Google Cloud Console → APIs & Services → Credentials.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = _run_local_server_interruptible(flow)
        TOKEN_FILE.write_text(creds.to_json())

    return build("calendar", "v3", credentials=creds)


# ── Local state helpers ───────────────────────────────────────────────────────

def load_state() -> dict:
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {"channels": []}

def save_state(state: dict):
    STATE_FILE.write_text(json.dumps(state, indent=2))

def add_channel(channel: dict):
    state = load_state()
    # Remove any existing entry with the same channel id (idempotent)
    state["channels"] = [c for c in state["channels"] if c.get("id") != channel.get("id")]
    state["channels"].append(channel)
    save_state(state)

def remove_channel(channel_id: str):
    state = load_state()
    state["channels"] = [c for c in state["channels"] if c.get("id") != channel_id]
    save_state(state)

def expiry_str(expiration_ms: str | None) -> str:
    if not expiration_ms:
        return "unknown"
    ts = int(expiration_ms) / 1000
    dt = datetime.fromtimestamp(ts, tz=timezone.utc)
    return dt.strftime("%Y-%m-%d %H:%M:%S UTC")

def is_expired(expiration_ms: str | None) -> bool:
    if not expiration_ms:
        return False
    ts = int(expiration_ms) / 1000
    return ts < datetime.now(tz=timezone.utc).timestamp()


# ── Commands ──────────────────────────────────────────────────────────────────

def cmd_calendars(args):
    """List all calendars accessible to the authenticated account."""
    service = get_service()
    try:
        result = service.calendarList().list().execute()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    items = result.get("items", [])
    if not items:
        print("No calendars found.")
        return

    for i, cal in enumerate(items):
        name = cal.get("summary", "?")
        cal_id = cal.get("id", "?")
        primary = " (primary)" if cal.get("primary") else ""
        print(f"  [{i+1}] {name}{primary}")
        print(f"      ID: {cal_id}")
    print()

def cmd_list(args):
    """List all locally tracked webhooks."""
    state = load_state()
    channels = state.get("channels", [])

    if not channels:
        print("No webhooks tracked locally. Use 'create' to add one.")
        return

    print(f"{'─'*70}")
    print(f"  {'Channel ID':<36}  {'Calendar':<20}  Status")
    print(f"{'─'*70}")
    for ch in channels:
        expired = is_expired(ch.get("expiration"))
        status = "⚠ EXPIRED" if expired else "✓ active"
        print(f"  {ch.get('id','?'):<36}  {ch.get('calendar_id','?'):<20}  {status}")
        print(f"    Resource ID : {ch.get('resourceId','?')}")
        print(f"    Webhook URL : {ch.get('address','?')}")
        print(f"    Expires     : {expiry_str(ch.get('expiration'))}")
        print()
    print(f"Total: {len(channels)} channel(s)")


def cmd_create(args):
    """Create a new webhook for a calendar."""
    service = get_service()

    channel_id = str(uuid.uuid4())
    body = {
        "id": channel_id,
        "type": "web_hook",
        "address": args.webhook_url,
        # Optional: token included in X-Goog-Channel-Token header for verification
        **({"token": args.token} if args.token else {}),
    }

    print(f"Creating webhook for calendar '{args.calendar_id}' → {args.webhook_url}")

    try:
        response = service.events().watch(calendarId=args.calendar_id, body=body).execute()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    channel = {
        "id": response["id"],
        "resourceId": response.get("resourceId"),
        "calendar_id": args.calendar_id,
        "address": args.webhook_url,
        "expiration": response.get("expiration"),
        "token": args.token,
        "created_at": datetime.now(tz=timezone.utc).isoformat(),
    }
    add_channel(channel)

    print("✓ Webhook created successfully!")
    print(f"  Channel ID  : {response['id']}")
    print(f"  Resource ID : {response.get('resourceId')}")
    print(f"  Expires     : {expiry_str(response.get('expiration'))}")
    print(f"State saved to {STATE_FILE}")


def cmd_refresh(args):
    """Delete the old channel and create a fresh one (Google has no update API)."""
    service = get_service()

    # Delete old channel first
    print(f"Stopping old channel {args.channel_id} ...")
    try:
        service.channels().stop(body={
            "id": args.channel_id,
            "resourceId": args.resource_id,
        }).execute()
        remove_channel(args.channel_id)
        print("  Old channel stopped.")
    except Exception as exc:
        print(f"  Warning: could not stop old channel: {exc}")

    # Create new one
    channel_id = str(uuid.uuid4())
    body = {
        "id": channel_id,
        "type": "web_hook",
        "address": args.webhook_url,
        **({"token": args.token} if args.token else {}),
    }

    print(f"Creating new webhook for calendar '{args.calendar_id}' → {args.webhook_url}")
    try:
        response = service.events().watch(calendarId=args.calendar_id, body=body).execute()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    channel = {
        "id": response["id"],
        "resourceId": response.get("resourceId"),
        "calendar_id": args.calendar_id,
        "address": args.webhook_url,
        "expiration": response.get("expiration"),
        "token": args.token,
        "created_at": datetime.now(tz=timezone.utc).isoformat(),
    }
    add_channel(channel)

    print("✓ Webhook refreshed!")
    print(f"  New Channel ID  : {response['id']}")
    print(f"  New Resource ID : {response.get('resourceId')}")
    print(f"  Expires         : {expiry_str(response.get('expiration'))}")


def cmd_delete(args):
    """Stop / delete a webhook channel."""
    service = get_service()

    # Look up resource_id from state if not provided
    resource_id = args.resource_id
    if not resource_id:
        state = load_state()
        for ch in state.get("channels", []):
            if ch.get("id") == args.channel_id:
                resource_id = ch.get("resourceId")
                break
        if not resource_id:
            print("ERROR: --resource-id is required (or run 'list' to find it).")
            sys.exit(1)

    print(f"Deleting channel {args.channel_id} ...")
    try:
        service.channels().stop(body={
            "id": args.channel_id,
            "resourceId": resource_id,
        }).execute()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    remove_channel(args.channel_id)
    print("✓ Channel deleted and removed from local state.")


def cmd_refresh_all(args):
    """Refresh all expired (or all) webhooks tracked locally."""
    state = load_state()
    channels = state.get("channels", [])

    if not channels:
        print("No channels tracked locally.")
        return

    to_refresh = [c for c in channels if args.all or is_expired(c.get("expiration"))]
    if not to_refresh:
        print("No expired channels found. Use --all to refresh everything.")
        return

    for ch in to_refresh:
        print(f"Refreshing channel {ch['id']} (calendar: {ch['calendar_id']}) ...")
        # Patch args for cmd_refresh
        sub = argparse.Namespace(
            channel_id=ch["id"],
            resource_id=ch.get("resourceId"),
            calendar_id=ch["calendar_id"],
            webhook_url=ch["address"],
            token=ch.get("token"),
        )
        cmd_refresh(sub)

    print(f"Done. Refreshed {len(to_refresh)} channel(s).")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Google Calendar Webhook Manager",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # list
    sub.add_parser("list", help="List tracked webhooks")

    # create
    p_create = sub.add_parser("create", help="Create a new webhook")
    p_create.add_argument("--calendar-id", default="primary", help="Calendar ID (default: primary)")
    p_create.add_argument("--webhook-url", required=True, help="HTTPS URL to receive notifications")
    p_create.add_argument("--token", default=None, help="Optional verification token")

    # refresh (single)
    p_refresh = sub.add_parser("refresh", help="Refresh (renew) a single webhook")
    p_refresh.add_argument("--channel-id", required=True, help="Existing channel ID")
    p_refresh.add_argument("--resource-id", required=True, help="Existing resource ID")
    p_refresh.add_argument("--calendar-id", default="primary", help="Calendar ID")
    p_refresh.add_argument("--webhook-url", required=True, help="Webhook URL for the new channel")
    p_refresh.add_argument("--token", default=None, help="Optional verification token")

    # refresh-all
    p_rall = sub.add_parser("refresh-all", help="Refresh all expired webhooks")
    p_rall.add_argument("--all", action="store_true", help="Refresh even non-expired channels")

    # delete
    p_delete = sub.add_parser("delete", help="Delete a webhook")
    p_delete.add_argument("--channel-id", required=True, help="Channel ID to delete")
    p_delete.add_argument("--resource-id", default=None, help="Resource ID (looked up from state if omitted)")

    sub.add_parser("calendars", help="List all available calendars and their IDs")

    args = parser.parse_args()

    dispatch = {
        "calendars": cmd_calendars,
        "list": cmd_list,
        "create": cmd_create,
        "refresh": cmd_refresh,
        "refresh-all": cmd_refresh_all,
        "delete": cmd_delete,
    }
    dispatch[args.command](args)


if __name__ == "__main__":
    main()
