# Bot Provisioner - Reference Guide

This repository contains complete documentation and reference implementations for building a bot provisioner application that uses OAuth 2.0 confidential client flow to create Microsoft Teams bots.

## Repository Structure

```
bot-provisioner-demo/
├── ARCHITECTURE.md           # Complete technical architecture and requirements
├── reference-backend/        # Reference backend implementation
│   ├── server.js            # Fully documented Express server
│   ├── package.json         # Dependencies
│   └── README.md            # Backend setup and concepts
├── reference-frontend/       # Reference frontend implementation
│   ├── index.html           # Main application page
│   ├── app.js               # Frontend logic with comments
│   ├── redirect.html        # OAuth callback handler
│   ├── admin-consent-callback.html  # Admin consent result
│   └── README.md            # Frontend setup and concepts
├── backend/                  # Working implementation (demo)
├── frontend/                 # Working implementation (demo)
└── [other files...]
```

## How to Use This Repository

### 1. Start with Architecture Documentation

Read **[ARCHITECTURE.md](./ARCHITECTURE.md)** first. It explains:
- System participants and their roles
- Requirements for bot provisioning
- Complete provisioning flow with sequence diagram
- Azure AD app registration steps
- Backend and frontend requirements

This document provides the conceptual foundation and architectural guidance.

### 2. Study the Reference Implementations

The `reference-backend/` and `reference-frontend/` folders contain clean, well-documented implementations that serve as both:
- **Working code** you can run and test
- **Documentation** explaining why and how things work

#### Reference Backend
- OAuth authorization code flow with MSAL
- Token acquisition and management
- Consent checking via `acquireTokenSilent()`
- Microsoft Graph and TDP API integration
- Error handling and session management

**Key file:** `reference-backend/server.js` (~750 lines with extensive comments)

#### Reference Frontend
- OAuth redirect handling
- Consent checking UI
- Bot provisioning orchestration
- Credentials display
- Teams installation deep link

**Key files:**
- `reference-frontend/index.html` - UI structure
- `reference-frontend/app.js` - Complete frontend logic

### 3. Run the Reference Implementation

Quick start:

```bash
# Backend
cd reference-backend
npm install
export CLIENT_ID=your-client-id
export CLIENT_SECRET=your-client-secret
npm start

# Frontend (new terminal)
cd reference-frontend
python3 -m http.server 8080
```

Visit http://localhost:8080 and follow the on-screen instructions.

See individual README files for detailed setup instructions.

### 4. Build Your Own Implementation

Use the reference implementations as a guide while building your own:

1. **Azure Setup**: Follow the steps in ARCHITECTURE.md to register your multi-tenant app
2. **Backend**: Implement the endpoints described in "Backend Implementation Requirements"
3. **Frontend**: Build the UI following "Frontend Implementation Requirements"
4. **Testing**: Test with your Azure AD app and different tenants

## Key Concepts Explained

### Confidential Client vs Public Client

| Aspect | Confidential Client | Public Client |
|--------|-------------------|---------------|
| **Secret Storage** | Server-side (secure) | Cannot store secrets |
| **Examples** | Web servers, backend APIs | SPAs, mobile apps, CLI tools |
| **OAuth Flow** | Authorization Code Flow | Device Code Flow, PKCE |
| **This Project** | ✓ Uses confidential client | N/A |

### Why User.Read First?

The app authenticates with `User.Read` scope initially because:
1. Any user can consent to `User.Read` (no admin required)
2. Allows users to sign in even if admin consent is missing
3. Provides refresh token for silent token acquisition
4. Clear separation between authentication and authorization

### How Silent Token Acquisition Works

```javascript
// Step 1: User authenticates with User.Read
const response = await cca.acquireTokenByCode({ scopes: ['User.Read'], ... });

// MSAL caches the refresh token internally

// Step 2: Later, get tokens for different resources silently
const graphToken = await cca.acquireTokenSilent({
    account: response.account,
    scopes: ['https://graph.microsoft.com/Application.ReadWrite.All']
});

const tdpToken = await cca.acquireTokenSilent({
    account: response.account,
    scopes: ['https://dev.teams.microsoft.com/AppDefinitions.ReadWrite']
});
```

No user interaction needed! Uses the cached refresh token.

### Multi-Tenant Architecture

The app is registered in **one** tenant (CWA's tenant) but can be used by users from **any** tenant:

1. Register as multi-tenant in Azure Portal
2. Use `/common` authority endpoint
3. Each tenant grants admin consent independently
4. User's tenant admin controls access

## Common Scenarios

### Scenario 1: First-Time User Setup

```
1. User visits application
2. Clicks "Check Scopes" → redirected to Azure AD
3. Authenticates with Microsoft account
4. Consent check shows missing scopes
5. User sends admin consent URL to their admin
6. Admin grants consent for tenant
7. User refreshes and sees "All scopes granted"
8. Proceeds to provision bot
```

### Scenario 2: Admin Denies Consent

```
1. Admin clicks consent URL
2. Clicks "Cancel" on consent screen
3. Redirected to admin-consent-callback.html with error
4. Users from that tenant cannot use the application
5. Admin must grant consent for app to work
```

### Scenario 3: Provisioning a Bot

```
1. User fills in bot name and endpoint
2. Clicks "Start Provisioning"
3. Backend creates AAD app → returns clientId
4. Backend generates secret → returns clientSecret
5. Backend creates Teams app → returns teamsAppId
6. Backend registers bot with Bot Framework
7. User sees credentials and Teams installation link
8. User copies credentials to .env file
9. User clicks link to install bot in Teams
```

## Security Best Practices

### Backend
- ✅ Store client secret in environment variables or Key Vault
- ✅ Use HTTPS in production
- ✅ Validate all inputs
- ✅ Implement rate limiting
- ✅ Use persistent session storage (Redis) in production
- ✅ Configure CORS to allow only trusted origins
- ✅ Rotate client secrets periodically

### Frontend
- ✅ Never expose client secret to browser
- ✅ Use HTTPS for OAuth redirects
- ✅ Validate OAuth state parameter
- ✅ Store minimal data in localStorage
- ✅ Handle token expiration gracefully
- ✅ Show clear error messages

## Troubleshooting

### "Admin consent required" message persists
- Verify admin clicked the consent URL and approved all permissions
- Check Azure Portal → Enterprise Applications → CWA app → Permissions
- Ensure admin consented for correct tenant
- Try clearing browser cache and re-authenticating

### OAuth redirect fails
- Verify redirect URIs in Azure Portal match exactly
- Check for trailing slashes (must match exactly)
- Ensure backend is running on expected port
- Check browser console for errors

### Bot provisioning fails at step X
- Check backend logs for detailed error messages
- Verify admin consent is granted (see above)
- Ensure user has Application Administrator role in Azure AD
- Check network connectivity to Microsoft APIs

### Teams installation link doesn't work
- Verify sideloading is enabled in tenant
- Check teamsAppId is correct
- Ensure bot endpoint is accessible via HTTPS
- Confirm user has permissions to install apps

## Additional Resources

- [Microsoft Identity Platform Documentation](https://docs.microsoft.com/azure/active-directory/develop/)
- [MSAL Node Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)
- [Microsoft Graph API Reference](https://docs.microsoft.com/graph/api/overview)
- [Teams Developer Portal API](https://dev.teams.microsoft.com)
- [Bot Framework Documentation](https://docs.microsoft.com/azure/bot-service/)

## Support

This is a reference implementation for educational purposes. For production deployments:
- Review and adapt code for your specific requirements
- Implement additional security measures
- Add monitoring and logging
- Follow your organization's security and compliance policies

## License

This reference implementation is provided as-is for educational purposes.
