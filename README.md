# calscripts

A collection of calendar scripts to help you stay organized. Access your calendars from the command line without needing desktop applications installed.

## Scripts

### outlook_web.py - Microsoft Outlook Calendar (Web Scraping)

Fetches your next 14 days of calendar events by automating the Outlook web interface using Playwright. **No admin consent required** - just uses your normal browser login.

**Features:**
- Uses your regular Microsoft login (no Azure app registration needed)
- Multiple account support via `config.toml` targets
- **CLI-only mode**: Headless browser with password prompt and 2FA code display
- Auto-login with credentials from config
- Output formats: text, iCal (.ics), JSON
- POST to URL with mTLS authentication
- Cross-platform: macOS, Linux (including WSL), Windows

### outlook.py - Microsoft Outlook Calendar (Graph API)

Alternative approach using the Microsoft Graph API. Requires Azure AD app registration and may need admin consent for `Calendars.Read` permission.

## Setup

### Prerequisites

1. **Python 3.13+**
2. **uv** package manager

### Installation

```bash
# Clone/download the project, then:
uv sync
```

### Playwright Browser Setup (for outlook_web.py)

After `uv sync`, install the browser(s):

```bash
# macOS - WebKit (Safari engine) is default and recommended
uv run playwright install webkit

# Linux/WSL/Windows - use Chromium
uv run playwright install chromium
```

**Note:** On Linux/WSL, use `--browser chromium` when running the script.

### Configuration (optional)

Copy `config.toml.example` to `config.toml` and configure your accounts:

```toml
# Target accounts
[targets.work]
username = "user@company.com"
# password = "..."  # Optional - omit for manual login

[targets.personal]
username = "user@outlook.com"

# POST endpoint (optional)
[post]
url = "https://api.example.com/calendar"

# mTLS certificates (optional - defaults to ~/.config/cauth/)
[mtls]
ca = "~/.config/cauth/ca.pem"
cert = "~/.config/cauth/crt.pem"
key = "~/.config/cauth/key.pem"
```

### Microsoft Azure App Registration (for outlook.py - Graph API)

Only needed if using the Graph API version. Requires Azure AD app registration and may need admin consent:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Set **Redirect URI** to `http://localhost:8000` (Web)
4. Add `Calendars.Read` permission under **API permissions** → **Microsoft Graph**
5. Edit `outlook.py` with your **Client ID** and **Tenant ID**

## Usage

### With browser GUI

```bash
uv run python outlook_web.py --target work
```

### CLI-only mode (no GUI)

Run completely from the command line with `--cli`:
- Prompts for password if not in config
- Displays 2FA verification number in terminal
- Runs browser headlessly (no window)

```bash
# Will prompt for password and show 2FA code
uv run python outlook_web.py --target work --cli

# With JSON output
uv run python outlook_web.py --target work --cli --json

# POST to endpoint
uv run python outlook_web.py --target work --cli --json --post
```

### Output formats

```bash
# Text output (default)
uv run python outlook_web.py -t work

# iCal format
uv run python outlook_web.py -t work --ical
uv run python outlook_web.py -t work --ical -o calendar.ics

# JSON format
uv run python outlook_web.py -t work --json
uv run python outlook_web.py -t work --json -o calendar.json
```

### POST to URL with mTLS

```bash
uv run python outlook_web.py -t work --json --post
```

### List configured targets

```bash
uv run python outlook_web.py --list-targets
```

### Alternative: Graph API

```bash
uv run python outlook.py
```

Requires Azure app registration (see setup above).

## Adding New Scripts

This project is designed to collect multiple calendar-related scripts. To add a new script:

1. Create a new Python file (e.g., `google_calendar.py`)
2. Add any new dependencies with `uv add <package>`
3. Optionally add a script entry to `pyproject.toml` under `[project.scripts]`
4. Update this README with documentation

## License

MIT