# Architecture

## Overview

A Node.js/Express backend that provisions Microsoft Teams bots using the Microsoft 365 Agents Toolkit's authentication approach. Uses MSAL for OAuth 2.0 Device Code Flow and acquires tokens for multiple resources (Microsoft Graph + Teams Dev Portal).

## Components

```
┌─────────────┐         ┌──────────────┐         ┌─────────────┐
│   Browser   │────────▶│   Express    │────────▶│   MSAL      │
│  (Frontend) │         │   Backend    │         │   Library   │
└─────────────┘         └──────────────┘         └─────────────┘
                               │                         │
                               │                         │
                               ▼                         ▼
                        ┌──────────────┐         ┌─────────────┐
                        │   Sessions   │         │  Azure AD   │
                        │  (In-Memory) │         │   Tokens    │
                        └──────────────┘         └─────────────┘
                               │
                               ▼
                    ┌─────────────────────┐
                    │   External APIs:    │
                    │  • Microsoft Graph  │
                    │  • Teams Dev Portal │
                    └─────────────────────┘
```

## Authentication Flow

### 1. Device Code Initiation
```javascript
POST /api/auth/start
```
- MSAL calls `acquireTokenByDeviceCode()` with Graph scopes
- Returns device code + user code + verification URL
- MSAL handles token acquisition in background
- Stores pending request with promise tracking

**Key:** MSAL automatically caches tokens for `acquireTokenSilent()` to work.

### 2. Polling for Completion
```javascript
POST /api/auth/poll
```
- Frontend polls every 3 seconds
- Backend checks if MSAL promise resolved
- When complete: extracts account info, creates session
- Session stores account object (not raw tokens)

### 3. Silent Token Acquisition
```javascript
await pca.acquireTokenSilent({
  account: session.account,
  scopes: ['https://dev.teams.microsoft.com/AppDefinitions.ReadWrite']
})
```
- MSAL uses cached refresh token
- Gets new access token for different resource
- No user interaction needed
- Works across resource servers (Graph → TDP)

## Provisioning Flow

```
1. Create AAD App          → Graph API
   GET token via acquireTokenSilent(graphScopes)
   POST /v1.0/applications

2. Generate Client Secret  → Graph API
   GET token via acquireTokenSilent(graphScopes)
   POST /v1.0/applications/{id}/addPassword

3. Create Teams App        → TDP API
   GET token via acquireTokenSilent(tdpScopes)  ← Different resource!
   POST /api/appdefinitions/v2/import

4. Register Bot            → TDP API
   GET token via acquireTokenSilent(tdpScopes)
   POST /api/botframework
```

## Key Technical Decisions

### Why MSAL's acquireTokenByDeviceCode()?
- Automatically caches tokens + refresh tokens
- Makes `acquireTokenSilent()` work seamlessly
- Manual token requests don't populate MSAL cache

### Why In-Memory Sessions?
- Simple for demo purposes
- Stores account objects (for MSAL), not tokens
- Production: use Redis or similar

### Why Two Token Requests?
OAuth 2.0 limitation: one access token per resource server.
- **Token 1:** `https://graph.microsoft.com` (for AAD operations)
- **Token 2:** `https://dev.teams.microsoft.com` (for Teams operations)

Both use same cached refresh token via `acquireTokenSilent()`.

### Why Device Code Flow?
- Perfect for CLI/desktop apps (no browser redirect needed)
- User authenticates on any device
- Matches Microsoft 365 Agents Toolkit CLI behavior

## Session Management

```javascript
sessions.set(sessionId, {
  account: {
    homeAccountId: '...',
    tenantId: '...',
    username: '...',
    // ... MSAL account object
  },
  createdAt: Date.now()
})
```

No raw tokens stored. MSAL manages tokens internally.

## API Endpoints

| Endpoint | Purpose | Auth Required |
|----------|---------|---------------|
| `POST /api/auth/start` | Start device code flow | No |
| `POST /api/auth/poll` | Check auth completion | No |
| `POST /api/provision/aad-app` | Create Azure AD app | Session |
| `POST /api/provision/client-secret` | Generate secret | Session |
| `POST /api/provision/teams-app` | Create Teams app | Session |
| `POST /api/provision/bot` | Register bot | Session |

All provisioning endpoints use `getTokenForScopes()` helper which calls `acquireTokenSilent()`.

## Security

- No secrets in code (only public CLI client ID)
- No tokens stored in session (MSAL manages them)
- CORS enabled for local development only
- Sessions stored in-memory (ephemeral)

## Differences from Real CLI

| CLI | This Demo |
|-----|-----------|
| File-based token cache (~/.fx) | In-memory sessions |
| Authorization Code Flow (browser) | Device Code Flow (any device) |
| Persistent across restarts | Ephemeral (restart clears state) |

Core MSAL usage is identical.
