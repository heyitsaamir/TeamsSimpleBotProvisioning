# Bot Provisioner - Reference Implementation

Documentation and reference code for building a confidential OAuth client that provisions Microsoft Teams bots.

## Documentation

**[ARCHITECTURE.md](./ARCHITECTURE.md)** - Complete technical guide covering:
- System architecture and requirements
- Azure AD app setup
- Backend and frontend implementation requirements
- Provisioning flow with sequence diagrams

## Reference Implementations

**[reference-backend/](./reference-backend/)** - Backend server with OAuth, consent checking, and bot provisioning

**[reference-frontend/](./reference-frontend/)** - Frontend with authentication flow and provisioning UI

Both implementations include extensive inline documentation.

## Quick Start

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

Visit http://localhost:8080

See individual README files in each folder for details.
