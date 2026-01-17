# Bot Provisioner - Reference Backend

This is a reference implementation of the backend server for the Bot Provisioner application. It demonstrates how to implement a confidential OAuth client that provisions Microsoft Teams bots.

## Setup

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure Environment Variables

Create a `.env` file or set these environment variables:

```bash
# Required: Your Azure AD app registration details
CLIENT_ID=your-client-id-here
CLIENT_SECRET=your-client-secret-here

# Optional: Customize redirect URIs (defaults to localhost:8080)
REDIRECT_URI=http://localhost:8080/redirect.html
ADMIN_CONSENT_URI=http://localhost:8080/admin-consent-callback.html

# Optional: Customize port (defaults to 3003)
PORT=3003
```

### 3. Run the Server

```bash
# Development (with auto-reload)
npm run dev

# Production
npm start
```

## Architecture

This server implements four main capabilities:

### 1. Authentication & Authorization
- **GET /api/auth/start** - Initiates OAuth flow with User.Read scope
- **POST /api/auth/callback** - Exchanges authorization code for tokens

### 2. Consent Checking
- **POST /api/auth/check-consent** - Verifies admin consent status for required scopes

### 3. Tenant Validation
- **POST /api/check-sideloading** - Checks if tenant allows custom app uploads

### 4. Bot Provisioning
- **POST /api/provision/aad-app** - Creates Azure AD app registration
- **POST /api/provision/client-secret** - Generates client secret
- **POST /api/provision/teams-app** - Creates Teams app package
- **POST /api/provision/bot** - Registers bot with Bot Framework

## Key Concepts

### Confidential Client
This server uses MSAL's `ConfidentialClientApplication`, which requires a client secret. The secret is stored server-side and never exposed to the browser.

### Silent Token Acquisition
After initial authentication, the server uses `acquireTokenSilent()` to get access tokens for different resources (Graph, TDP) without requiring user interaction. This works because MSAL caches the refresh token.

### Multi-Tenant
The app is registered in one Azure AD tenant but can authenticate users from any tenant. This is enabled by:
- Registering as multi-tenant in Azure Portal
- Using `/common` authority endpoint
- Each tenant must grant admin consent separately

### Error Handling
The server distinguishes between expected consent errors (`consent_required`, `invalid_grant` with AADSTS65001) and unexpected errors (network issues, token expiration). This allows the frontend to show appropriate UI for missing consent.

## Code Structure

The code is organized into clear sections with extensive comments:

1. **Configuration** - App settings and API endpoints
2. **MSAL Setup** - Confidential client initialization
3. **Authentication Endpoints** - OAuth flow handlers
4. **Scope Checking** - Consent verification
5. **Sideloading Check** - Tenant settings validation
6. **Provisioning Endpoints** - Bot creation flow
7. **Helper Functions** - Reusable utilities

Each function includes detailed comments explaining:
- What it does
- Why it's needed
- How it works
- Key concepts to understand

## Security Notes

- Never commit `.env` files or secrets to version control
- Use environment variables or Azure Key Vault for secrets in production
- Sessions are stored in-memory - use Redis or similar for production
- CORS is enabled for all origins - restrict in production
- Client secret should be rotated periodically

## Testing

You can test endpoints using tools like curl or Postman:

```bash
# Start OAuth flow
curl http://localhost:3003/api/auth/start

# Check consent (requires valid sessionId)
curl -X POST http://localhost:3003/api/auth/check-consent \
  -H "Content-Type: application/json" \
  -d '{"sessionId": "your-session-id"}'
```

## Production Considerations

Before deploying to production:

1. **Session Storage**: Replace in-memory Map with Redis
2. **CORS**: Configure allowed origins explicitly
3. **Logging**: Add structured logging (e.g., Winston, Application Insights)
4. **Error Handling**: Add more specific error messages and status codes
5. **Rate Limiting**: Protect against abuse
6. **HTTPS**: Always use HTTPS in production
7. **Secret Management**: Use Azure Key Vault or similar
8. **Monitoring**: Add health checks and metrics
