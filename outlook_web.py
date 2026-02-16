#!/usr/bin/env python3
"""
Fetch Outlook calendar events via web automation (Playwright).
No admin consent required - uses your normal browser login.

Supports:
- Multiple target accounts via config.toml
- Auto-login with credentials from config
- Output as text, iCal, or JSON
- POST to URL with mTLS authentication

Works on: macOS, Linux (including WSL), Windows
"""

import argparse
import hashlib
import json
import os
import re
import shutil
import ssl
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import httpx
import tomli
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# Paths
BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.toml"
DEFAULT_MTLS_DIR = Path.home() / ".config" / "cauth"


def load_config() -> dict:
    """Load configuration from config.toml."""
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "rb") as f:
            return tomli.load(f)
    return {}


def get_credentials(config: dict, target: str) -> tuple[str | None, str | None]:
    """
    Get username and password for a target from config.
    """
    targets = config.get("targets", {})
    if target not in targets:
        return None, None
    
    target_config = targets[target]
    username = target_config.get("username")
    password = target_config.get("password")
    
    if username and not password:
        print(f"⚠ No password in config for {username} - manual login required", file=sys.stderr)
    
    return username, password


def get_browser_data_dir(browser: str) -> Path:
    """Get browser-specific data directory."""
    return BASE_DIR / f".browser_data_{browser}"

def get_calendar_events(
    days: int = 14,
    headless: bool = False,
    browser: str = "webkit",
    username: str | None = None,
    password: str | None = None
):
    """
    Fetch calendar events from Outlook Web.
    
    Args:
        days: Number of days to look ahead
        headless: Run browser without GUI (set False for first login)
        browser: Browser to use - 'webkit' (Safari) or 'chromium' (Chrome)
        username: Optional username for auto-login
        password: Optional password for auto-login
    """
    events = []
    user_data_dir = get_browser_data_dir(browser)
    
    with sync_playwright() as p:
        # Select browser engine
        if browser == "webkit":
            browser_type = p.webkit
        else:
            browser_type = p.chromium
        
        # Use persistent context to remember login
        context = browser_type.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            headless=headless,
            viewport={"width": 1280, "height": 900},
            slow_mo=100,  # Slow down actions slightly for stability
        )
        
        page = context.pages[0] if context.pages else context.new_page()
        
        # Calculate date range
        start_date = datetime.now()
        end_date = start_date + timedelta(days=days)
        
        calendar_url = "https://outlook.office.com/calendar/view/month"
        print(f"Opening Outlook calendar (using {browser})...", file=sys.stderr)
        page.goto(calendar_url, wait_until="domcontentloaded", timeout=60000)
        
        # Wait a moment for any redirects
        page.wait_for_timeout(2000)
        
        # Check if we need to log in - check current URL
        current_url = page.url
        needs_login = any(x in current_url for x in [
            "login.microsoftonline.com",
            "login.live.com", 
            "login.microsoft.com",
            "microsoftonline.com/oauth",
            "microsoftonline.com/common"
        ])
        
        if needs_login:
            # Try auto-login if credentials provided
            if username:
                print(f"Attempting auto-login for {username}...", file=sys.stderr)
                try:
                    # Wait for email input field
                    email_input = page.locator('input[type="email"], input[name="loginfmt"]')
                    email_input.wait_for(timeout=10000)
                    email_input.fill(username)
                    
                    # Click next
                    next_btn = page.locator('input[type="submit"], button[type="submit"]')
                    next_btn.click()
                    page.wait_for_timeout(2000)
                    
                    # If password provided, fill it in
                    if password:
                        password_input = page.locator('input[type="password"], input[name="passwd"]')
                        password_input.wait_for(timeout=10000)
                        password_input.fill(password)
                        
                        # Click sign in
                        signin_btn = page.locator('input[type="submit"], button[type="submit"]')
                        signin_btn.click()
                        page.wait_for_timeout(3000)
                        
                        # Handle "Stay signed in?" prompt if it appears
                        try:
                            no_btn = page.locator('text=/No|Decline/i').first
                            if no_btn:
                                no_btn.click(timeout=3000)
                        except:
                            pass
                except Exception as e:
                    print(f"Auto-login failed: {e}", file=sys.stderr)
                    print("Falling back to manual login...", file=sys.stderr)
            
            # Check if still needs login
            page.wait_for_timeout(2000)
            current_url = page.url
            still_needs_login = any(x in current_url for x in [
                "login.microsoftonline.com",
                "login.live.com",
                "login.microsoft.com"
            ])
            
            if still_needs_login:
                print("\n" + "=" * 50, file=sys.stderr)
                print("🔐 LOGIN REQUIRED", file=sys.stderr)
                print("=" * 50, file=sys.stderr)
                print("Please complete login in the browser (MFA may be required).", file=sys.stderr)
                print("Take your time - the script will wait up to 10 minutes.", file=sys.stderr)
                print("=" * 50 + "\n", file=sys.stderr)
            
            # Wait for redirect back to calendar (up to 10 minutes for login + MFA)
            try:
                page.wait_for_url(
                    lambda url: "outlook.office.com/calendar" in url or "outlook.office365.com/calendar" in url,
                    timeout=600000  # 10 minutes
                )
                print("\n✓ Login successful!\n", file=sys.stderr)
            except PlaywrightTimeout:
                print("\n❌ Login timeout. Please try again.", file=sys.stderr)
                context.close()
                return []
        
        # Wait for calendar to fully load
        print("Waiting for calendar to load...", file=sys.stderr)
        page.wait_for_load_state("networkidle", timeout=30000)
        page.wait_for_timeout(3000)  # Extra wait for dynamic content
        
        # Switch to agenda/list view for easier scraping
        # Try to click on "List" or "Agenda" view if available
        try:
            # Look for view switcher
            view_button = page.locator('[aria-label*="view"]').first
            if view_button:
                view_button.click()
                page.wait_for_timeout(500)
                
                # Try to find and click "Agenda" or "List" option
                list_option = page.locator('text=/^(Agenda|List)$/i').first
                if list_option:
                    list_option.click()
                    page.wait_for_timeout(1000)
        except:
            pass  # Continue with current view
        
        # Extract events using JavaScript
        print(f"Extracting events for the next {days} days...", file=sys.stderr)
        
        # Method 1: Try to get events from the page's React state or data attributes
        events_data = page.evaluate("""
            () => {
                const events = [];
                
                // Find all calendar event elements
                const eventElements = document.querySelectorAll('[data-is-focusable="true"][role="button"], .ms-Callout-main, [class*="event"], [class*="calendar-item"]');
                
                eventElements.forEach(el => {
                    const text = el.innerText || el.textContent || '';
                    if (text.trim() && text.length > 2 && text.length < 500) {
                        // Try to extract time and title
                        const lines = text.split('\\n').map(l => l.trim()).filter(l => l);
                        if (lines.length > 0) {
                            events.push({
                                raw: text.trim(),
                                lines: lines
                            });
                        }
                    }
                });
                
                // Also try aria-label which often contains full event info
                document.querySelectorAll('[aria-label]').forEach(el => {
                    const label = el.getAttribute('aria-label');
                    if (label && label.includes(':') && (label.includes('AM') || label.includes('PM') || label.includes('meeting') || label.includes('event'))) {
                        events.push({
                            raw: label,
                            lines: label.split(',').map(l => l.trim())
                        });
                    }
                });
                
                return events;
            }
        """)
        
        # Alternative: Get the full page content and look for calendar data
        if not events_data:
            # Try getting from network requests or embedded JSON
            events_data = page.evaluate("""
                () => {
                    // Look for embedded calendar data
                    const scripts = document.querySelectorAll('script');
                    for (const script of scripts) {
                        const text = script.textContent || '';
                        if (text.includes('calendarEvents') || text.includes('appointments')) {
                            return { embedded: text.substring(0, 5000) };
                        }
                    }
                    return [];
                }
            """)
        
        # Sign out of Outlook to allow switching accounts
        print("Signing out of Outlook...", file=sys.stderr)
        try:
            # Navigate to Microsoft sign-out URL
            page.goto("https://login.microsoftonline.com/common/oauth2/v2.0/logout", timeout=15000)
            page.wait_for_timeout(2000)
            
            # Also clear cookies for complete logout
            context.clear_cookies()
            print("✓ Signed out successfully", file=sys.stderr)
        except Exception as e:
            # Don't fail if logout has issues - we still got the data
            print(f"Note: Logout may be incomplete: {e}", file=sys.stderr)
        
        context.close()
        
        # Remove the browser data directory to ensure clean state for next login
        if user_data_dir.exists():
            try:
                shutil.rmtree(user_data_dir)
            except:
                pass  # Ignore cleanup errors
        
        # Parse and deduplicate events
        seen = set()
        for item in events_data:
            if isinstance(item, dict):
                raw = item.get('raw', '')
                if raw and raw not in seen and len(raw) > 5:
                    seen.add(raw)
                    events.append(raw)
        
        return events


def parse_event(raw: str) -> dict | None:
    """
    Parse a raw event string into structured data.
    
    Expected format from aria-label:
    "Event Title, 10:00 AM to 11:00 AM, Tuesday, February 3, 2026, ..."
    or for all-day:
    "Event Title, all day event, Tuesday, February 3, 2026, ..."
    """
    parts = [p.strip() for p in raw.split(',')]
    if len(parts) < 3:
        return None
    
    title = parts[0]
    
    # Skip non-event items
    if title.startswith('calendar view') or title.startswith('current time'):
        return None
    
    # Find time and date parts
    time_part = None
    date_part = None
    all_day = False
    
    for i, part in enumerate(parts[1:], 1):
        # Check for time range like "10:00 AM to 11:00 AM"
        if ' to ' in part and ('AM' in part or 'PM' in part):
            time_part = part
        # Check for all-day
        elif 'all day' in part.lower():
            all_day = True
        # Check for date like "Tuesday, February 3, 2026"
        elif re.match(r'(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)', part):
            # Date spans multiple comma-separated parts
            # e.g., "Tuesday", "February 3", "2026"
            if i + 2 < len(parts):
                date_part = f"{part}, {parts[i+1]}, {parts[i+2]}"
            elif i + 1 < len(parts):
                date_part = f"{part}, {parts[i+1]}"
            else:
                date_part = part
            break
    
    if not date_part:
        return None
    
    # Parse the date and time
    try:
        # Try parsing date like "Tuesday, February 3, 2026"
        date_match = re.search(r'(\w+day),?\s+(\w+)\s+(\d+),?\s+(\d{4})', date_part)
        if not date_match:
            return None
        
        month_name = date_match.group(2)
        day = int(date_match.group(3))
        year = int(date_match.group(4))
        
        month_map = {
            'January': 1, 'February': 2, 'March': 3, 'April': 4,
            'May': 5, 'June': 6, 'July': 7, 'August': 8,
            'September': 9, 'October': 10, 'November': 11, 'December': 12
        }
        month = month_map.get(month_name, 1)
        
        if all_day:
            start_dt = datetime(year, month, day, 0, 0)
            end_dt = datetime(year, month, day, 23, 59)
        elif time_part:
            # Parse time like "10:00 AM to 11:00 AM"
            time_match = re.match(r'(\d{1,2}):(\d{2})\s*(AM|PM)\s+to\s+(\d{1,2}):(\d{2})\s*(AM|PM)', time_part)
            if time_match:
                start_hour = int(time_match.group(1))
                start_min = int(time_match.group(2))
                start_ampm = time_match.group(3)
                end_hour = int(time_match.group(4))
                end_min = int(time_match.group(5))
                end_ampm = time_match.group(6)
                
                if start_ampm == 'PM' and start_hour != 12:
                    start_hour += 12
                elif start_ampm == 'AM' and start_hour == 12:
                    start_hour = 0
                    
                if end_ampm == 'PM' and end_hour != 12:
                    end_hour += 12
                elif end_ampm == 'AM' and end_hour == 12:
                    end_hour = 0
                
                start_dt = datetime(year, month, day, start_hour, start_min)
                end_dt = datetime(year, month, day, end_hour, end_min)
            else:
                return None
        else:
            return None
        
        return {
            'title': title,
            'start': start_dt,
            'end': end_dt,
            'all_day': all_day,
            'raw': raw
        }
    except (ValueError, AttributeError):
        return None


def events_to_ical(events: list[dict]) -> str:
    """
    Convert parsed events to iCal format.
    """
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//calscripts//outlook_web//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]
    
    for event in events:
        # Generate a unique ID based on event content
        uid = hashlib.md5(f"{event['title']}{event['start']}".encode()).hexdigest()
        
        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}@calscripts")
        lines.append(f"DTSTAMP:{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}")
        
        if event['all_day']:
            lines.append(f"DTSTART;VALUE=DATE:{event['start'].strftime('%Y%m%d')}")
            lines.append(f"DTEND;VALUE=DATE:{event['end'].strftime('%Y%m%d')}")
        else:
            lines.append(f"DTSTART:{event['start'].strftime('%Y%m%dT%H%M%S')}")
            lines.append(f"DTEND:{event['end'].strftime('%Y%m%dT%H%M%S')}")
        
        # Escape special characters in title
        title = event['title'].replace('\\', '\\\\').replace(',', '\\,').replace(';', '\\;')
        lines.append(f"SUMMARY:{title}")
        lines.append("END:VEVENT")
    
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)


def events_to_json(events: list[dict], target: str | None = None) -> str:
    """
    Convert parsed events to JSON format.
    """
    json_events = []
    for event in events:
        json_events.append({
            "title": event["title"],
            "start": event["start"].isoformat(),
            "end": event["end"].isoformat(),
            "all_day": event["all_day"]
        })
    
    output = {
        "target": target,
        "fetched_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "event_count": len(json_events),
        "events": json_events
    }
    return json.dumps(output, indent=2)


def post_to_url(data: str, url: str, config: dict) -> bool:
    """
    POST data to URL using mTLS.
    
    Args:
        data: JSON data to POST
        url: Target URL
        config: Configuration dict with mtls settings
    """
    # Get mTLS cert paths
    mtls_config = config.get("mtls", {})
    ca_path = Path(mtls_config.get("ca", DEFAULT_MTLS_DIR / "ca.pem")).expanduser()
    cert_path = Path(mtls_config.get("cert", DEFAULT_MTLS_DIR / "crt.pem")).expanduser()
    key_path = Path(mtls_config.get("key", DEFAULT_MTLS_DIR / "key.pem")).expanduser()
    
    # Verify cert files exist
    for path, name in [(ca_path, "CA"), (cert_path, "cert"), (key_path, "key")]:
        if not path.exists():
            print(f"❌ mTLS {name} file not found: {path}", file=sys.stderr)
            return False
    
    try:
        # Create SSL context with system CAs (to verify server cert) + custom CA
        # This handles the case where server cert is signed by public CA (e.g. Let's Encrypt)
        # but mTLS requires our custom client cert
        ssl_context = ssl.create_default_context()
        
        # Add our custom CA (in case server cert is signed by it)
        ssl_context.load_verify_locations(cafile=str(ca_path))
        
        # Load client certificate and key for mTLS authentication
        ssl_context.load_cert_chain(certfile=str(cert_path), keyfile=str(key_path))
        
        with httpx.Client(
            verify=ssl_context,
            timeout=30.0
        ) as client:
            response = client.post(
                url,
                content=data,
                headers={"Content-Type": "application/json"}
            )
            response.raise_for_status()
            print(f"✓ Posted to {url} (status: {response.status_code})", file=sys.stderr)
            return True
            
    except httpx.HTTPStatusError as e:
        print(f"❌ HTTP error: {e.response.status_code} - {e.response.text}", file=sys.stderr)
    except Exception as e:
        print(f"❌ POST failed: {e}", file=sys.stderr)
    
    return False


def main():
    parser = argparse.ArgumentParser(
        description="Fetch Outlook calendar events via web automation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Fetch calendar with manual login
  %(prog)s
  
  # Use a configured target account
  %(prog)s --target work
  
  # Output as JSON and POST to URL
  %(prog)s --target work --json --post
  
  # Save iCal to file
  %(prog)s --target personal --ical -o calendar.ics
"""
    )
    parser.add_argument(
        "--target", "-t",
        type=str,
        help="Target account name from config.toml"
    )
    parser.add_argument(
        "--browser", "-b",
        choices=["webkit", "chromium"],
        default="webkit",
        help="Browser engine: 'webkit' (Safari) or 'chromium' (Chrome). Default: webkit"
    )
    parser.add_argument(
        "--days", "-d",
        type=int,
        default=14,
        help="Number of days to look ahead. Default: 14"
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run in headless mode (no GUI). Only use after first login."
    )
    parser.add_argument(
        "--ical", "-i",
        action="store_true",
        help="Output in iCal format (.ics)"
    )
    parser.add_argument(
        "--json", "-j",
        action="store_true",
        help="Output in JSON format"
    )
    parser.add_argument(
        "--post", "-p",
        action="store_true",
        help="POST JSON output to URL configured in config.toml (requires --json)"
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        help="Write output to file (e.g., calendar.ics or calendar.json)"
    )
    parser.add_argument(
        "--list-targets",
        action="store_true",
        help="List available targets from config.toml and exit"
    )
    args = parser.parse_args()
    
    # Load configuration
    config = load_config()
    
    # List targets and exit
    if args.list_targets:
        targets = config.get("targets", {})
        if targets:
            print("Available targets:")
            for name, cfg in targets.items():
                print(f"  {name}: {cfg.get('username', '(no username)')}")
        else:
            print("No targets configured. Create config.toml from config.toml.example")
        return
    
    # Get credentials if target specified
    username, password = None, None
    if args.target:
        username, password = get_credentials(config, args.target)
        if not username:
            print(f"❌ Target '{args.target}' not found in config.toml", file=sys.stderr)
            print("Use --list-targets to see available targets", file=sys.stderr)
            return
    
    # Validate --post requires --json
    if args.post and not args.json:
        print("❌ --post requires --json flag", file=sys.stderr)
        return
    
    # Suppress status messages if outputting to stdout
    quiet = (args.ical or args.json) and not args.output and not args.post
    
    if not quiet:
        print("📅 Outlook Web Calendar Fetcher", file=sys.stderr)
        print("=" * 40, file=sys.stderr)
        if args.target:
            print(f"Target: {args.target} ({username})", file=sys.stderr)
    
    raw_events = get_calendar_events(
        days=args.days,
        headless=args.headless,
        browser=args.browser,
        username=username,
        password=password
    )
    
    if not raw_events:
        print("\n⚠️  No events found. The page structure may have changed.", file=sys.stderr)
        print("   Try running again - the calendar might need more time to load.", file=sys.stderr)
        return
    
    # Parse events into structured format
    parsed_events = []
    for raw in raw_events:
        event = parse_event(raw)
        if event:
            parsed_events.append(event)
    
    if not quiet:
        print(f"\n📋 Found {len(parsed_events)} calendar events\n", file=sys.stderr)
    
    # Output based on format
    if args.json:
        json_output = events_to_json(parsed_events, target=args.target)
        
        if args.post:
            post_url = config.get("post", {}).get("url")
            if not post_url:
                print("❌ No POST URL configured in config.toml", file=sys.stderr)
                return
            post_to_url(json_output, post_url, config)
        
        if args.output:
            with open(args.output, 'w') as f:
                f.write(json_output)
            print(f"✓ Saved to {args.output}", file=sys.stderr)
        elif not args.post:
            print(json_output)
            
    elif args.ical:
        ical_output = events_to_ical(parsed_events)
        
        if args.output:
            with open(args.output, 'w') as f:
                f.write(ical_output)
            print(f"✓ Saved to {args.output}", file=sys.stderr)
        else:
            print(ical_output)
    else:
        # Text output
        for event in parsed_events:
            start_str = event['start'].strftime('%a %b %d %H:%M')
            end_str = event['end'].strftime('%H:%M')
            if event['all_day']:
                print(f"  {event['start'].strftime('%a %b %d')} (all day) - {event['title']}")
            else:
                print(f"  {start_str}-{end_str} - {event['title']}")


if __name__ == "__main__":
    main()
