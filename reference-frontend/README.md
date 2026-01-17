# Bot Provisioner - Reference Frontend

This is a reference implementation of the frontend for the Bot Provisioner application. It demonstrates how to build a web interface for OAuth authentication, consent checking, and bot provisioning.

## Overview

The frontend consists of three main pages:

1. **index.html** - Main application page with consent checking and provisioning UI
2. **redirect.html** - OAuth callback handler (where Azure AD redirects after login)
3. **admin-consent-callback.html** - Admin consent callback handler

## Setup

### Prerequisites

You need a way to serve static files. Choose one:

#### Option 1: Python HTTP Server (Simplest)
```bash
# Python 3
python3 -m http.server 8080

# Python 2
python -m SimpleHTTPServer 8080
```

#### Option 2: Node.js HTTP Server
```bash
npm install -g http-server
http-server -p 8080
```

#### Option 3: VS Code Live Server Extension
1. Install "Live Server" extension in VS Code
2. Right-click `index.html` and select "Open with Live Server"

### Configuration

Edit `app.js` and `redirect.html` to point to your backend:

```javascript
const API_BASE = 'http://localhost:3003/api';  // Change this to match your backend URL
```

If your backend is running on a different port or domain, update this value.

## Usage

### Step 1: Check Scopes
1. Open `http://localhost:8080/index.html` in your browser
2. Click "Check Scopes"
3. You'll be redirected to Azure AD to sign in
4. After signing in, you'll be redirected back to the app
5. The app checks if your tenant admin has granted required permissions

### Step 2: Admin Consent (if needed)
If scopes are missing:
1. Copy the admin consent URL displayed on the page
2. Send it to your tenant administrator
3. Admin clicks the URL and grants consent
4. Admin is redirected to the consent callback page
5. Return to main page and click "Check Scopes" again

### Step 3: Check Sideloading (Optional)
1. Click "Check Sideloading" to verify your tenant allows custom apps
2. If disabled, contact your Teams administrator

### Step 4: Provision Bot
1. Fill in bot name and endpoint URL
2. Click "Start Provisioning"
3. Wait for provisioning to complete (creates AAD app, secret, Teams app, bot)
4. Copy the generated credentials to your `.env` file
5. Click the Teams installation link to install the bot

## Architecture

### Authentication Flow

```
User clicks "Check Scopes"
    ↓
app.js calls GET /api/auth/start
    ↓
Backend returns Azure AD authorization URL
    ↓
User redirected to Azure AD (login page)
    ↓
User authenticates with Microsoft account
    ↓
Azure AD redirects to redirect.html with authorization code
    ↓
redirect.html calls POST /api/auth/callback with code
    ↓
Backend exchanges code for tokens, returns sessionId
    ↓
redirect.html stores sessionId in localStorage
    ↓
Redirects back to index.html
    ↓
app.js uses sessionId for subsequent API calls
```

### Provisioning Flow

```
User clicks "Start Provisioning"
    ↓
app.js calls backend endpoints sequentially:
    1. POST /api/provision/aad-app → creates AAD app
    2. POST /api/provision/client-secret → generates secret
    3. POST /api/provision/teams-app → creates Teams app
    4. POST /api/provision/bot → registers bot
    ↓
Display credentials and Teams installation link
```

## Key Concepts

### Session Management
- **sessionId** is stored in localStorage for persistence across page reloads
- If session expires (401/403), user must re-authenticate
- Session only contains account info, not tokens (tokens are managed by backend)

### OAuth Redirect Flow
1. User clicks button → redirected to Azure AD
2. Azure AD redirects to `redirect.html` with code
3. `redirect.html` exchanges code for tokens via backend
4. User redirected back to main page with session established

### Admin Consent URL
The admin consent URL format is:
```
https://login.microsoftonline.com/{tenantId}/adminconsent
  ?client_id={clientId}
  &redirect_uri={adminConsentRedirectUri}
```

This grants tenant-wide consent for all requested permissions.

### Teams Deep Link
After provisioning, the frontend generates a Teams installation link:
```
https://teams.microsoft.com/l/app/{teamsAppId}
  ?installAppPackage=true
  &webjoin=true
  &appTenantId={tenantId}
  &login_hint={userPrincipalName}
```

This allows one-click installation of the bot in Teams.

## Code Structure

### index.html
- Main application UI with three steps (consent check, provision, results)
- Uses CSS for styling (inline for simplicity)
- Minimal, functional design focused on clarity

### app.js
The JavaScript is organized into clear sections:

1. **Configuration** - API endpoint configuration
2. **Initialization** - Session restoration and event listeners
3. **Authentication & Scope Checking** - OAuth flow and consent verification
4. **Sideloading Check** - Tenant settings validation
5. **Bot Provisioning** - Complete provisioning flow orchestration
6. **Results Display** - Credentials and installation link

Each function includes comments explaining:
- What it does
- How it works
- What it returns
- Error handling

### redirect.html
- Minimal page that processes OAuth callback
- Extracts authorization code from URL
- Calls backend to exchange code for tokens
- Stores session and redirects to main page
- Handles authentication errors gracefully

### admin-consent-callback.html
- Simple page confirming consent grant or denial
- Explains what the consent means
- Provides link back to main application
- No auto-redirect (user controls when to return)

## Security Considerations

- **No secrets in frontend** - Client secret is never exposed to browser
- **HTTPS required** - OAuth requires HTTPS in production (localhost HTTP is OK for dev)
- **State parameter** - Used for CSRF protection in OAuth flow
- **sessionId** - Random identifier, not predictable
- **localStorage** - Contains no sensitive data (just sessionId and username)

## Production Considerations

Before deploying to production:

1. **Use HTTPS** - Required for OAuth in production
2. **Configure CORS** - Backend must allow your frontend domain
3. **Update redirect URIs** - Register production URLs in Azure AD app
4. **Error handling** - Add more user-friendly error messages
5. **Loading states** - Add better visual feedback for long operations
6. **Validation** - Add client-side validation for bot endpoint format
7. **Styling** - Enhance UI/UX for better user experience
8. **Analytics** - Track user flows and errors

## Troubleshooting

### "Invalid session" error
- Session expired or backend restarted
- Click "Check Scopes" to re-authenticate

### OAuth redirect not working
- Check redirect URI matches exactly in Azure AD app registration
- Ensure backend is running on correct port
- Check browser console for errors

### Provisioning fails
- Verify admin consent has been granted
- Check backend logs for detailed error messages
- Ensure bot endpoint is valid HTTPS URL

### Teams installation link doesn't work
- Verify sideloading is enabled in tenant
- Check teamsAppId is correct
- Ensure user has permissions to install apps in Teams
