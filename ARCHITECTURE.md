# Bot Provisioner - Technical Architecture

## System Participants

### CodeWriterApp (CWA)
The confidential client application that orchestrates bot provisioning. Key characteristics:
- **Backend**: Server component that handles OAuth flows and API calls to Graph/TDP
- **Frontend**: User interface for authentication, consent checking, and bot provisioning
- **Registered in**: CWA's own Azure AD tenant as a multi-tenant application
- **Authentication**: Uses OAuth 2.0 authorization code flow with client secret (confidential client)

### Azure AD (CWA Tenant)
The identity provider where CWA is registered as a multi-tenant application. Handles:
- OAuth authorization code flow for user authentication
- Token issuance for accessing Microsoft Graph and Teams Dev Portal
- Admin consent grants (tenant-wide permissions)

### User (User Tenant)
An end user from any Microsoft 365 tenant who wants to provision a bot. The user:
- Authenticates via OAuth to CWA
- Must have sufficient permissions to create apps in their tenant
- May need their tenant admin to grant consent for required permissions

### Admin (User Tenant)
A tenant administrator in the same tenant as the User. The admin:
- Has permissions to grant tenant-wide consent for applications
- Receives admin consent URL from User when permissions are needed
- Grants consent for CWA to access Graph and TDP APIs on behalf of users in their tenant

### Teams Developer Portal Backend (TDP)
Microsoft's backend service for Teams app management (`https://dev.teams.microsoft.com`). Provides APIs for:
- Creating Teams app definitions
- Registering bot endpoints
- Checking tenant sideloading status
- **Requires**: `AppDefinitions.ReadWrite` scope

### Microsoft Graph Backend
Microsoft's unified API endpoint (`https://graph.microsoft.com`). Provides APIs for:
- Creating Azure AD app registrations
- Generating client secrets
- Managing app permissions
- **Requires**: `Application.ReadWrite.All` scope

## Requirements for Bot Provisioning

For CWA to successfully provision a Teams bot for a user, the following conditions must be met:

### 1. User Authentication
- User must authenticate with CWA using OAuth 2.0 authorization code flow
- Minimum scope: `User.Read` (always grantable by user)

### 2. Admin Consent
The user's tenant admin must grant consent for these delegated permissions:
- **Graph API**: `Application.ReadWrite.All` - Create Azure AD apps and secrets
- **Teams Dev Portal**: `AppDefinitions.ReadWrite` - Create Teams apps and register bots

### 3. User Permissions
The authenticated user must have roles in their tenant allowing:
- Creating app registrations in Azure AD (typically requires Application Administrator role)
- Creating Teams apps via Developer Portal

### 4. Tenant Settings
- **Sideloading**: Tenant must allow custom app uploads (verifiable via TDP API)
- **App Registration**: Tenant must allow users to register applications (Azure AD setting)

### 5. Bot Endpoint
- User must provide an HTTPS endpoint where the bot will be hosted
- Endpoint must be accessible from Microsoft Teams infrastructure

## Provisioning Flow

```mermaid
sequenceDiagram
    participant Admin
    participant User
    participant CWA Frontend
    participant CWA Backend
    participant Azure AD
    participant Graph API
    participant TDP API

    Note over Admin,TDP API: Initial Authentication & Consent Check

    User->>CWA Frontend: Begin Scope Check
    CWA Frontend->>CWA Backend: GET /api/auth/start
    CWA Backend->>Azure AD: Generate auth URL (User.Read)
    Azure AD-->>CWA Backend: Authorization URL
    CWA Backend-->>CWA Frontend: Return auth URL
    CWA Frontend->>Azure AD: Redirect user to auth URL (on AAD)
    Azure AD->>User: Show Microsoft login
    User->>Azure AD: Authenticate + consent to User.Read
    Azure AD->>CWA Frontend: Redirect with auth code
    CWA Frontend->>CWA Backend: POST /api/auth/callback (code)
    CWA Backend->>Azure AD: Exchange code for tokens
    Azure AD-->>CWA Backend: Access token + refresh token

    Note over Admin,TDP API: Scope Verification

    loop For each required scope
        CWA Backend->>Azure AD: acquireTokenSilent(scope)
        alt Scope granted
            Azure AD-->>CWA Backend: Access token
        else Scope missing
            Azure AD-->>CWA Backend: consent_required error
        end
    end
    CWA Backend-->>CWA Frontend: Granted + missing scopes

    alt Missing scopes
        CWA Frontend->>User: Display admin consent URL
        User-->>Admin: Informs Admin and gives them URL
        Admin->>Azure AD: Admin clicks consent URL
        Azure AD->>Admin: Show admin consent screen
        Admin->>Azure AD: Admin grants tenant-wide consent for CWA AAD App
        Azure AD->>CWA Frontend: Redirect to consent callback
        CWA Frontend->>Admin: Show "Consent granted" message
        Admin-->>User: Informs user that their request was granted
        User->>CWA Frontend: Begin scope check again
        Note over CWA Backend,Azure AD: Now acquireTokenSilent succeeds
    end

    Note over Admin,TDP API: Sideloading Check

    User->>CWA Frontend: Click "Check Sideloading"
    CWA Frontend->>CWA Backend: POST /api/check-sideloading (sessionId)
    CWA Backend->>Azure AD: acquireTokenSilent(TDP scope)
    Azure AD-->>CWA Backend: TDP access token
    CWA Backend->>TDP API: GET /api/usersettings/mtUserAppPolicy
    TDP API-->>CWA Backend: isSideloadingAllowed: true/false
    CWA Backend-->>CWA Frontend: Sideloading status
    CWA Frontend->>User: Display sideloading status

    Note over Admin,TDP API: Bot Provisioning

    User->>CWA Frontend: Assuming Bot details (name, endpoint, app package) are available<br/>Start Bot creation

    Note over CWA Backend,Graph API: Create Azure AD App Registration
    CWA Frontend->>CWA Backend: POST /api/provision/aad-app<br/>(sessionId, appName)
    CWA Backend->>Azure AD: acquireTokenSilent(Graph scope)
    Azure AD-->>CWA Backend: Graph access token
    CWA Backend->>Graph API: POST /applications<br/>(displayName, signInAudience)
    Graph API-->>CWA Backend: App created (clientId, appRegistrationId)

    Note over CWA Backend,Graph API: Generate Client Secret
    CWA Backend->>Graph API: POST /applications/{appRegistrationId}/addPassword<br/>(displayName, endDateTime)
    Graph API-->>CWA Backend: Client secret generated (clientSecret)

    Note over CWA Backend,TDP API: Create Teams App Package
    CWA Backend->>Azure AD: acquireTokenSilent(TDP scope)
    Azure AD-->>CWA Backend: TDP access token
    CWA Backend->>TDP API: POST /api/appdefinitions/v2/import<br/>(app package zip: manifest.json, color.png, outline.png)
    TDP API-->>CWA Backend: Teams app created (teamsAppId)

    Note over CWA Backend,TDP API: Register Bot with Bot Framework
    CWA Backend->>TDP API: POST /api/botframework<br/>(botId, name, messagingEndpoint, configuredChannels)
    TDP API-->>CWA Backend: Bot registered
    CWA Backend-->>CWA Frontend: Success

    CWA Frontend->>User: Display credentials (clientId, clientSecret, teamsAppId)
```

## Key Technical Details

### Token Acquisition Strategy
CWA uses a two-phase token acquisition approach:

1. **Initial auth**: Request only `User.Read` scope - this always succeeds, allowing users to sign in even without admin consent
2. **Subsequent requests**: Use `acquireTokenSilent()` with cached refresh token to obtain tokens for admin-consented scopes

This pattern avoids blocking users at sign-in while clearly indicating which permissions are missing.

### Multi-Resource Token Management
The provisioning flow requires tokens for two different resource servers:
- **Microsoft Graph**: For Azure AD app management
- **Teams Dev Portal**: For Teams app and bot management

Azure AD does not allow requesting scopes from multiple resources in a single auth request. CWA handles this by:
- Authenticating once with `User.Read`
- Using the cached refresh token to silently acquire tokens for each resource as needed

### Error Handling
When `acquireTokenSilent()` fails, CWA distinguishes between:
- **Expected**: `consent_required`, `interaction_required`, `invalid_grant` with AADSTS65001 → Show admin consent URL
- **Unexpected**: Network errors, token expiration, unknown errors → Return error to user

### Session Management
- Sessions stored in-memory on backend (Map keyed by sessionId)
- Frontend stores sessionId and userInfo in localStorage
- MSAL handles token caching internally
- Sessions lost on backend restart (in-memory only)
