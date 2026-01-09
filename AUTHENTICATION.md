# Authentication Implementation

## How the Microsoft 365 Agents Toolkit CLI Authenticates

The actual Microsoft 365 Agents Toolkit CLI uses a specific Azure AD application for authentication that has the necessary permissions pre-configured by Microsoft.

### Client ID Used by CLI

The CLI uses: `7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0`

This can be verified in the source code:
- `/packages/cli/src/commonlib/m365Login.ts` (line 37)
- `/packages/cli/src/commonlib/common/userPasswordConfig.ts` (line 10)
- `/packages/cli/src/commonlib/azureLogin.ts` (line 70)

## Key Improvement: Single Authentication with acquireTokenSilent()

This demo now uses **ONE authentication** just like the real CLI, instead of requiring two separate logins.

### Why the Teams Toolkit Client ID Didn't Work

Initially, this demo used the Teams Toolkit public client ID: `7ab7862c-4c57-491e-8a45-d52a7e023983`

This caused the following error:
```
AADSTS65002: Consent between first party application '7ab7862c-4c57-491e-8a45-d52a7e023983'
and first party resource '00000003-0000-0000-c000-000000000000' must be configured via
preauthorization - applications owned and operated by Microsoft must get approval from
the API owner before requesting tokens for that API.
```

**Explanation:**
- Both the Teams Toolkit client and Microsoft Graph are "first-party" (Microsoft-owned) applications
- Microsoft restricts which of their own apps can access certain Graph API scopes
- The Teams Toolkit client ID doesn't have pre-authorization for `Application.ReadWrite.All`
- The actual CLI client ID does have the necessary pre-authorization

### Authentication Flow

The demo uses **OAuth 2.0 Device Code Flow**, which is perfect for CLI/desktop applications:

1. **Request Device Code (ONE TIME)**
   ```
   POST https://login.microsoftonline.com/common/oauth2/v2.0/devicecode
   ```
   - Requests Graph API scopes ONLY:
     - `https://graph.microsoft.com/Application.ReadWrite.All`
     - `https://graph.microsoft.com/TeamsAppInstallation.ReadForUser`
   - **Important**: Cannot mix scopes from different resource servers! Each token is for ONE resource.
   - Teams Dev Portal scopes (`https://dev.teams.microsoft.com/*`) will be acquired later
   - Returns a user code and device code
   - User visits Microsoft login page and enters the code

2. **Poll for Token**
   ```
   POST https://login.microsoftonline.com/common/oauth2/v2.0/token
   ```
   - Polls Azure AD to check if user has completed authentication
   - Returns access token + ID token when complete
   - **Caches the account information** from ID token

3. **Silent Token Acquisition (For Different Resources)**
   - Uses `pca.acquireTokenSilent()` from `@azure/msal-node`
   - Gets tokens for specific scopes without re-prompting user
   - When requesting scopes for a **different resource** (like TDP after Graph):
     - MSAL uses the cached refresh token
     - Makes token request to Azure AD for the new resource
     - User may see incremental consent prompt (but no re-login)
     - Returns token for the new resource
   - Example for Graph API:
     ```javascript
     const token = await pca.acquireTokenSilent({
       account: cachedAccount,
       scopes: ['https://graph.microsoft.com/Application.ReadWrite.All'],
       forceRefresh: false
     });
     ```
   - Example for Teams Dev Portal:
     ```javascript
     const token = await pca.acquireTokenSilent({
       account: cachedAccount,
       scopes: ['https://dev.teams.microsoft.com/AppDefinitions.ReadWrite'],
       forceRefresh: false
     });
     ```

### Session Management

The backend maintains sessions with cached account information (not raw tokens):
```javascript
sessions.set(sessionId, {
  account: {
    homeAccountId: `${oid}.${tid}`,
    environment: 'login.microsoftonline.com',
    tenantId: tid,
    username: username,
    localAccountId: oid,
    name: name,
    idTokenClaims: { ... }
  },
  createdAt: Date.now(),
});
```

This cached account is then used with `acquireTokenSilent()` to get tokens for specific scopes on-demand, without storing raw tokens or requiring re-authentication.

## Key OAuth/Azure AD Concept: One Token Per Resource

**Important**: Each access token is for **ONE resource server** only. You cannot mix scopes from different resources in a single token request.

### Resource Servers

- **Microsoft Graph**: `https://graph.microsoft.com`
  - Scopes: `https://graph.microsoft.com/Application.ReadWrite.All`, etc.
- **Teams Dev Portal**: `https://dev.teams.microsoft.com`
  - Scopes: `https://dev.teams.microsoft.com/AppDefinitions.ReadWrite`

### Why Two Separate Token Requests

1. **Initial Authentication**: Request Graph API scopes
   - User authenticates and consents
   - Azure AD returns Graph API token
   - Account + refresh token cached by MSAL

2. **Later (When Needed)**: Request TDP scopes via `acquireTokenSilent()`
   - MSAL uses cached refresh token
   - Azure AD returns TDP token (different resource!)
   - User may see consent prompt but no re-login required

This is **not two authentications** - it's one authentication, then silent token acquisition for additional resources.

### How to Use This Demo

1. Start the backend:
   ```bash
   cd backend
   npm install
   node server.js
   ```

2. Open the frontend:
   ```bash
   cd frontend
   python3 -m http.server 8080
   # Then visit http://localhost:8080
   ```

3. Follow the wizard:
   - **Step 1**: Authenticate with Microsoft 365 (Graph API scopes)
   - **Step 2**: Configure bot name and endpoint
   - **Step 3**: Provisioning happens automatically
     - First 2 operations use Graph API token
     - When Teams app creation starts, `acquireTokenSilent()` gets TDP token
     - You may briefly see a consent prompt for TDP scope (no re-login needed)
   - **Step 4**: Download credentials

Note: The consent screen will show Graph API permissions initially. When provisioning reaches Teams app creation, you may see another quick consent for Teams Dev Portal access - this is normal and demonstrates how `acquireTokenSilent()` handles multiple resources!

### What Gets Created

After successful provisioning, you'll have:
- **Azure AD Application**: With client ID and secret
- **Teams App**: Registered in Teams Dev Portal with manifest
- **Bot Registration**: Connected to Bot Framework with messaging endpoint

### Credentials Generated

- `CLIENT_ID`: Azure AD app client ID (also used as BOT_ID)
- `CLIENT_SECRET`: Password credential for the app
- `TENANT_ID`: Your Azure AD tenant ID
- `TEAMS_APP_ID`: The Teams app manifest ID
- `BOT_ENDPOINT`: Your bot's messaging endpoint

These credentials can be used directly with the Bot Framework SDK to run your bot.
