# Confidential Client Implementation

This directory contains an alternative implementation using **OAuth 2.0 Authorization Code Flow** with a **Confidential Client** instead of the device code flow with a public client.

## Key Differences from Public Client

| Aspect | Public Client | Confidential Client |
|--------|---------------|---------------------|
| **Flow** | Device Code Flow | Authorization Code Flow + PKCE |
| **MSAL Type** | `PublicClientApplication` | `ConfidentialClientApplication` |
| **Client Secret** | Not required | **Required** |
| **User Experience** | Copy/paste device code | Browser redirect (seamless) |
| **Port** | 3001 | 3002 |
| **Files** | `server.js`, `index.html` | `server-confidential.js`, `index-confidential.html`, `redirect.html` |

## Setup

### 1. Azure AD App Configuration

Your Azure AD app (`2a098349-9ecc-463f-a053-d5675e10deeb`) must be configured as:

**Authentication:**
- Platform: **Web** (not "Mobile and desktop applications")
- Redirect URI: `http://localhost:8080/redirect.html`
- For production: `https://your-domain.com/redirect.html`

**Certificates & secrets:**
- Create a **client secret**
- Copy the secret value (you'll only see it once!)

**API permissions:**
- Microsoft Graph (Delegated):
  - `Application.ReadWrite.All`
  - `TeamsAppInstallation.ReadForUser`
- Microsoft Teams (search for "Microsoft Teams" in "APIs my org uses"):
  - `AppDefinitions.ReadWrite` (Delegated)

**CRITICAL: Admin Consent Required**

⚠️ Azure AD **does NOT allow requesting scopes from multiple resources** in a single authorization request. Since Graph API and Teams Dev Portal are different resource servers, we can only request Graph scopes during initial user auth.

For Teams Dev Portal scopes, you **MUST** have an admin pre-consent:

1. Go to Azure Portal → App registrations → Your app
2. **API permissions** → **Add a permission**
3. **APIs my organization uses** → Search for **"Microsoft Teams"**
4. Select **"Microsoft Teams"** (Application ID: `2cd2e3b4-31db-4288-90f0-800084e46016`)
5. **Delegated permissions** → Check **`AppDefinitions.ReadWrite`**
6. Click **Add permissions**
7. Click **"Grant admin consent for [Your Tenant]"** ← **This is critical!**

Without admin pre-consent, bot provisioning will fail when trying to create the Teams app.

### 2. Environment Variable

Set your client secret as an environment variable:

```bash
export CLIENT_SECRET="your-secret-value-here"
```

Or create a `.env` file in the `backend` directory:
```
CLIENT_SECRET=your-secret-value-here
```

### 3. Start the Server

```bash
cd backend
CLIENT_SECRET=your-secret node server-confidential.js
```

The server will start on **port 3002**.

### 4. Start the Frontend

```bash
cd frontend
python3 -m http.server 8080
```

Open `http://localhost:8080/index-confidential.html`

## How It Works

### Authentication Flow

1. **User clicks "Sign in with Microsoft"**
   - Frontend calls `GET /api/auth/start`
   - Backend generates PKCE code verifier and challenge
   - Backend returns authorization URL
   - Frontend redirects to Microsoft login

2. **User authenticates with Microsoft**
   - Microsoft redirects back to `http://localhost:8080/redirect.html?code=xxx&state=yyy`

3. **Redirect page handles callback**
   - Extracts `code` and `state` from URL
   - Calls `POST /api/auth/callback` with code and state
   - Backend exchanges code for tokens using `acquireTokenByCode()`
   - MSAL caches tokens and account

4. **Redirect page stores session**
   - Stores `sessionId` and `userInfo` in `localStorage`
   - Redirects to `index-confidential.html`

5. **Provisioning continues normally**
   - Uses `acquireTokenSilent()` for additional scopes (same as public client)

## Architecture

```
┌─────────────────┐
│   Frontend      │
│ (port 8080)     │
└────────┬────────┘
         │
         │ 1. GET /api/auth/start
         │    Returns: authUrl
         ▼
┌─────────────────┐
│   Backend       │      ┌──────────────┐
│ (port 3002)     │──────│     MSAL     │
│                 │      │ Confidential │
│ + Client Secret │      │    Client    │
└────────┬────────┘      └──────────────┘
         │
         │ 2. User redirected to Microsoft
         │
         ▼
┌─────────────────────────┐
│  Microsoft Login        │
│  (login.microsoft.com)  │
└────────┬────────────────┘
         │
         │ 3. Redirect to /redirect.html?code=xxx
         │
         ▼
┌─────────────────┐
│  redirect.html  │
│                 │
│ POST /api/auth/callback
│   { code, state }
└─────────────────┘
         │
         │ 4. Backend exchanges code for tokens
         │    Returns: sessionId, userInfo
         │
         ▼
┌─────────────────┐
│ index-          │
│ confidential.   │
│ html            │
│                 │
│ (Provisioning   │
│  continues)     │
└─────────────────┘
```

## Security Features

1. **PKCE (Proof Key for Code Exchange)**
   - Generates random code verifier
   - Creates SHA256 hash as code challenge
   - Prevents authorization code interception attacks

2. **State Parameter**
   - Random value for CSRF protection
   - Validated on callback to ensure request originated from your app

3. **Client Secret**
   - Stored server-side only
   - Never exposed to frontend/browser
   - Validates server identity to Microsoft

## Production Considerations

1. **HTTPS Required**
   - Microsoft requires HTTPS redirect URIs in production
   - Update `redirectUri` in `server-confidential.js`
   - Update redirect URI in Azure AD app registration

2. **Secure Secret Storage**
   - Use Azure Key Vault, AWS Secrets Manager, or similar
   - Never commit secrets to git
   - Rotate secrets periodically

3. **Session Storage**
   - Replace in-memory `Map` with Redis or similar
   - Add session expiration
   - Implement session cleanup

4. **Error Handling**
   - Add proper logging
   - Handle token refresh failures
   - Implement retry logic

## Comparison: When to Use Each Flow

### Use Public Client (Device Code Flow)
- ✅ CLI tools, terminal applications
- ✅ Headless servers, IoT devices
- ✅ No web server required
- ✅ Better for scripting/automation
- ✅ No client secret management
- ❌ Worse UX (copy/paste code)
- ❌ User must manually navigate to URL

### Use Confidential Client (Authorization Code Flow)
- ✅ Web applications with backend
- ✅ Better UX (seamless redirect)
- ✅ More secure (client authenticated)
- ✅ Standard OAuth flow
- ❌ Requires client secret
- ❌ Requires web server
- ❌ More complex setup

## Troubleshooting

### "Client secret not configured"
```bash
# Set environment variable before starting server
export CLIENT_SECRET="your-secret"
node server-confidential.js
```

### "Redirect URI mismatch"
- Check Azure AD app registration
- Ensure redirect URI exactly matches: `http://localhost:8080/redirect.html`
- Check for trailing slashes

### "Invalid state parameter"
- State may have expired (15 minute timeout)
- Try authenticating again

### "CORS errors"
- Frontend must be on port 8080
- Backend must be on port 3002
- Check CORS configuration in server-confidential.js

## Files

| File | Purpose |
|------|---------|
| `backend/server-confidential.js` | Express server with confidential client |
| `frontend/index-confidential.html` | Main UI for confidential client version |
| `frontend/app-confidential.js` | Frontend logic for confidential client |
| `frontend/redirect.html` | OAuth callback handler page |

## Testing

1. Start backend: `CLIENT_SECRET=xxx node backend/server-confidential.js`
2. Start frontend: `python3 -m http.server 8080` (from frontend directory)
3. Open: `http://localhost:8080/index-confidential.html`
4. Click "Sign in with Microsoft"
5. Should redirect to Microsoft login
6. After login, should redirect back to `redirect.html`
7. Should then redirect to `index-confidential.html` with session established

## Next Steps

- Add proper error boundaries in frontend
- Implement token refresh logic
- Add session cleanup/expiration
- Move to production redirect URIs (HTTPS)
- Store client secret in secure vault
- Add logging and monitoring
