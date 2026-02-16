#!/usr/bin/env python3
import msal
import requests
from datetime import datetime, timedelta
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse

# === CONFIGURE THESE ===
CLIENT_ID = "0e6a228a-3017-462e-aad4-28af9b6f9129"
TENANT_ID = "b1519f0f-2dbf-4e21-bf34-a686ce97588a"
CLIENT_SECRET = "your-client-secret"  # Optional for public client flow
# =======================

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Calendars.Read"]
REDIRECT_URI = "http://localhost:8000"

auth_code = None

class AuthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        query = urllib.parse.urlparse(self.path).query
        params = urllib.parse.parse_qs(query)
        auth_code = params.get("code", [None])[0]
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"<h1>Login successful! You can close this window.</h1>")

    def log_message(self, format, *args):
        pass

def get_access_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    
    # Try cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    
    # Interactive login
    auth_url = app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    print("Opening browser for login...")
    webbrowser.open(auth_url)
    
    server = HTTPServer(("localhost", 8000), AuthHandler)
    server.handle_request()
    
    if auth_code:
        result = app.acquire_token_by_authorization_code(auth_code, SCOPES, redirect_uri=REDIRECT_URI)
        if "access_token" in result:
            return result["access_token"]
    
    raise Exception("Failed to get access token")

def get_calendar_events(token):
    now = datetime.utcnow()
    end = now + timedelta(days=14)
    
    url = "https://graph.microsoft.com/v1.0/me/calendarView"
    headers = {"Authorization": f"Bearer {token}"}
    params = {
        "startDateTime": now.isoformat() + "Z",
        "endDateTime": end.isoformat() + "Z",
        "$orderby": "start/dateTime",
        "$select": "subject,start,end,location"
    }
    
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    return response.json().get("value", [])

def main():
    token = get_access_token()
    events = get_calendar_events(token)
    
    print(f"\n📅 Calendar events for the next 14 days ({len(events)} found):\n")
    for event in events:
        start = datetime.fromisoformat(event["start"]["dateTime"].replace("Z", ""))
        subject = event.get("subject", "(No subject)")
        location = event.get("location", {}).get("displayName", "")
        loc_str = f" @ {location}" if location else ""
        print(f"  {start.strftime('%a %b %d %H:%M')} - {subject}{loc_str}")

if __name__ == "__main__":
    main()
