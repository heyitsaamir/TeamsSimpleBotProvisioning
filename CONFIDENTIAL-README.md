# Bot Provisioner - Confidential Client

Simplified implementation that uses confidential client for auth/consent checking only, then falls back to proven provisioning logic.

## Architecture

### Backend (server-confidential.js)
1. **OAuth with User.Read** - User signs in with minimal scope (always works, no admin blocking)
2. **Check Service Principal** - Queries Graph API to detect missing permissions
3. **Provisioning** - Reuses all endpoints from original server.js

### Frontend
1. **admin-consent-callback.html** - Simple callback that says "The client is successfully consented for your tenant"
2. **index-confidential.html** - Basic HTML (no styling), shows admin consent URL if needed
3. **app-confidential.js** - Handles OAuth redirect and provisioning

## Quick Start

1. **Set environment variables:**
   ```bash
   export CLIENT_SECRET=your_secret_here
   export ENDPOINT_BASE=https://your-tunnel-url
   ```

2. **Run the start script:**
   ```bash
   ./start-confidential.sh
   ```

This will:
- Start backend on `http://localhost:3003`
- Start frontend on `http://localhost:8080`
- Open browser automatically

## Configuration

Update redirect URIs in Azure AD app registration to match your `ENDPOINT_BASE`:
- `${ENDPOINT_BASE}/redirect.html`
- `${ENDPOINT_BASE}/admin-consent-callback.html`

Then access via your tunnel:
```
${ENDPOINT_BASE}/index-confidential.html
```

## How It Works

### Step 1: User Authentication
- User clicks "Authenticate"
- Redirects to Microsoft login with `User.Read` scope only
- User can always sign in (no admin blocking)

### Step 2: Consent Check
- Backend attempts to acquire tokens for required scopes
- Checks if these scopes are granted:
  - `Application.ReadWrite.All` (Graph)
  - `AppDefinitions.ReadWrite` (Teams Dev Portal)

### Step 3a: If Admin Consent Missing
- Shows warning with admin consent URL
- User copies URL and sends to admin
- Admin consents for entire tenant
- User refreshes page

### Step 3b: If Admin Consent Granted
- Step 2 enabled immediately
- User can provision bots

### Step 4: Provisioning
- Uses same proven logic from original server.js
- Creates AAD app, generates secret, creates Teams app, registers bot

## Files

### Backend
- `backend/server-confidential.js` - Port 3003

### Frontend
- `frontend/index-confidential.html` - Main page
- `frontend/app-confidential.js` - Frontend logic
- `frontend/redirect.html` - OAuth callback
- `frontend/admin-consent-callback.html` - Admin consent callback

### Scripts
- `start.sh` - Start both servers

## Environment Variables

- `CLIENT_SECRET` - Azure AD client secret (required)
- `ENDPOINT_BASE` - Base URL for redirect URIs (default: `https://3hvfdfhp-8080.usw2.devtunnels.ms`)

## Ports

- Backend: 3003
- Frontend: 8080

## Key Features

Features:
- Confidential client only used for auth/consent checking
- Provisioning reuses proven public client logic from server.js
- Basic HTML with no styling
- Admin consent callback just says "success" and redirects
