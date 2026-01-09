# Teams Bot Provisioner

Standalone web app for provisioning Microsoft Teams bots. Uses the same approach as the Microsoft 365 Agents Toolkit CLI.

## Quick Start

```bash
# Install and start backend
cd backend
npm install
node server.js

# Open frontend (in new terminal)
cd frontend
python3 -m http.server 8080
# Visit http://localhost:8080
```

Or use the quick start script:
```bash
./start.sh
```

## What It Does

Provisions a Teams bot in 4 steps:

1. **Authenticate** - OAuth device code flow (like `az login`)
2. **Configure** - Enter bot name and endpoint URL
3. **Provision** - Creates:
   - Azure AD app + client secret
   - Teams app registration
   - Bot Framework registration
4. **Download** - Get credentials as `.env` file

## Tech Stack

**Backend:**
- Express.js server (port 3001)
- MSAL for authentication
- Calls Microsoft Graph + Teams Dev Portal APIs

**Frontend:**
- Vanilla HTML/JS (no framework)
- 4-step wizard UI

## How It Works

See [ARCHITECTURE.md](./ARCHITECTURE.md) for details on:
- MSAL device code flow
- `acquireTokenSilent()` for multiple resources
- Session management
- Provisioning flow

## Requirements

- Node.js 16+
- Microsoft 365 account with admin privileges
- Admin consent for `Application.ReadWrite.All` scope

## Output

After provisioning, you get:

```env
CLIENT_ID=xxx
CLIENT_SECRET=xxx
TENANT_ID=xxx
TEAMS_APP_ID=xxx
BOT_ENDPOINT=https://your-bot.com
```

Plus a Teams deep link to install the app.

## Security Notes

⚠️ **This is a demo.** For production:
- Use Redis for sessions (not in-memory)
- Add HTTPS
- Restrict CORS
- Add rate limiting

