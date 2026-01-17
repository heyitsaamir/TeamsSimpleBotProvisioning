# Bot Provisioner - Documentation & Reference Implementation

This repository contains complete technical documentation and reference implementations for building a confidential OAuth client application that provisions Microsoft Teams bots.

## What's Included

### 📖 Documentation
- **[ARCHITECTURE.md](./ARCHITECTURE.md)** - Complete technical architecture guide
  - System participants and requirements
  - Provisioning flow with sequence diagrams
  - Azure AD app setup instructions
  - Backend and frontend implementation requirements

### 💻 Reference Implementations
- **[reference-backend/](./reference-backend/)** - Backend server with extensive inline documentation
  - OAuth authorization code flow with MSAL confidential client
  - Scope checking and admin consent handling
  - Microsoft Graph and Teams Developer Portal API integration
  - Complete bot provisioning workflow

- **[reference-frontend/](./reference-frontend/)** - Frontend application with detailed comments
  - OAuth redirect handling
  - Consent checking UI
  - Bot provisioning orchestration
  - Credentials display and Teams installation links

## Quick Start

### 1. Read the Architecture Documentation

Start with **[ARCHITECTURE.md](./ARCHITECTURE.md)** to understand:
- How the system works
- Requirements for bot provisioning
- OAuth flows and token management
- Setup instructions for Azure AD

### 2. Run the Reference Implementation

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

See individual README files in each folder for detailed setup instructions.

### 3. Build Your Own

Use the reference implementations as a guide:
1. Register your multi-tenant app in Azure AD (see ARCHITECTURE.md)
2. Implement backend endpoints following `reference-backend/server.js`
3. Build frontend UI following `reference-frontend/` structure
4. Test with your Azure AD app

## Key Features

### Confidential Client OAuth Flow
- Uses OAuth 2.0 authorization code flow with client secret
- Multi-tenant: works across any Azure AD tenant
- Admin consent workflow for required permissions
- Silent token acquisition for Graph and TDP APIs

### Bot Provisioning
Creates a complete Teams bot by:
1. Creating Azure AD app registration
2. Generating client secret
3. Creating Teams app package
4. Registering bot with Bot Framework

### Educational Focus
All code includes extensive comments explaining:
- **What** it does
- **Why** it's needed
- **How** it works
- Key concepts and decisions

## Architecture Overview

```
┌─────────────┐         ┌──────────────┐         ┌─────────────┐
│    User     │────────▶│     CWA      │────────▶│  Azure AD   │
│  (Browser)  │         │   Frontend   │         │   (OAuth)   │
└─────────────┘         └──────────────┘         └─────────────┘
                               │
                               ▼
                        ┌──────────────┐         ┌─────────────┐
                        │     CWA      │────────▶│    MSAL     │
                        │   Backend    │         │ (Token Mgmt)│
                        └──────────────┘         └─────────────┘
                               │
                    ┌──────────┴──────────┐
                    ▼                     ▼
            ┌──────────────┐      ┌──────────────┐
            │ Microsoft    │      │    Teams     │
            │   Graph      │      │ Dev Portal   │
            └──────────────┘      └──────────────┘
```

## Requirements

- **Azure AD App**: Multi-tenant app registration with confidential client
- **Node.js**: Version 16 or higher
- **Microsoft 365**: Account with Application Administrator role
- **Admin Consent**: For `Application.ReadWrite.All` and `AppDefinitions.ReadWrite` scopes

## What You'll Learn

- OAuth 2.0 authorization code flow for confidential clients
- Multi-tenant application architecture
- Admin consent and tenant-wide permissions
- Silent token acquisition with MSAL
- Multi-resource token management (Graph + TDP)
- Microsoft Graph API for Azure AD management
- Teams Developer Portal API for bot creation

## Repository Structure

```
bot-provisioner-demo/
├── README.md                 # This file
├── ARCHITECTURE.md           # Technical architecture guide
├── reference-backend/        # Backend implementation + docs
│   ├── server.js            # Fully documented server
│   ├── package.json         # Dependencies
│   └── README.md            # Backend setup guide
├── reference-frontend/       # Frontend implementation + docs
│   ├── index.html           # Main application page
│   ├── app.js               # Frontend logic with comments
│   ├── redirect.html        # OAuth callback handler
│   ├── admin-consent-callback.html
│   └── README.md            # Frontend setup guide
└── archived/                 # Original implementations
```

## Use Cases

This repository is intended for:
- **Learning** how to build confidential OAuth clients
- **Understanding** multi-tenant Azure AD applications
- **Reference** when implementing your own bot provisioner
- **Teaching** OAuth flows and Microsoft Graph integration

## Security Notes

The reference implementations demonstrate security best practices:
- Client secret stored server-side (never exposed to browser)
- Session-based state management
- CSRF protection with state parameter
- Clear separation of authentication and authorization
- Proper error handling for consent scenarios

**Production Checklist:**
- Use persistent session storage (Redis)
- Implement HTTPS everywhere
- Configure CORS for specific origins
- Add rate limiting
- Use Azure Key Vault for secrets
- Implement logging and monitoring
- Add health checks

## License

This is a reference implementation for educational purposes.

## Contributing

This repository serves as documentation and reference. For issues or improvements, please file an issue.
