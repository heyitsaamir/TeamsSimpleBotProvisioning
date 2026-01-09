# 🤖 Bot Provisioner Demo

A complete wizard-style web application for provisioning Microsoft Teams bots without the Microsoft 365 Agents Toolkit CLI.

## Features

✅ **Step-by-step wizard interface**
✅ **Device code authentication flow**
✅ **Real-time provisioning progress**
✅ **Credential management and export**
✅ **Zero framework dependencies** (vanilla HTML/CSS/JS frontend)
✅ **Simple Express.js backend**

## Architecture

```
┌─────────────────┐          ┌─────────────────┐          ┌──────────────────┐
│   Browser       │          │  Express.js     │          │  Microsoft       │
│   (Frontend)    │◄────────►│  Backend        │◄────────►│  365 APIs        │
│                 │   HTTP   │                 │   HTTPS  │                  │
│  - Wizard UI    │          │  - Auth         │          │  - Graph API     │
│  - Forms        │          │  - Provisioning │          │  - Teams Dev     │
│  - Progress     │          │  - Session Mgmt │          │    Portal API    │
└─────────────────┘          └─────────────────┘          └──────────────────┘
```

## Prerequisites

- **Node.js** 16+ and npm
- **Microsoft 365 account** with admin privileges
- **Admin consent** for `Application.ReadWrite.All` scope

## Quick Start

### 1. Install Backend Dependencies

```bash
cd backend
npm install
```

### 2. Start Backend Server

```bash
npm start
```

The backend will start on `http://localhost:3001`

### 3. Open Frontend

```bash
cd ../frontend
# Open in browser (no build step needed!)
open index.html
# Or use a simple HTTP server:
python3 -m http.server 8080
# Then visit: http://localhost:8080
```

### 4. Follow the Wizard

The wizard will guide you through:
1. **Authentication** - Sign in with your M365 account
2. **Configuration** - Enter bot name and endpoint
3. **Provisioning** - Automated resource creation
4. **Credentials** - Download your bot credentials

## Usage Guide

### Step 1: Authentication

1. Click "Authenticate with Microsoft Graph"
2. A device code will be displayed
3. Open the link and enter the code
4. Sign in with your Microsoft 365 account
5. Grant the requested permissions
6. Repeat for "Authenticate with Teams Dev Portal"

**Required Permissions:**
- `Application.ReadWrite.All` (requires admin consent)
- `TeamsAppInstallation.ReadForUser`
- `AppDefinitions.ReadWrite`

### Step 2: Configure Bot

Fill in the form:

| Field | Description | Example |
|-------|-------------|---------|
| **Bot Name** | Display name for your bot | `Customer Support Bot` |
| **Bot Endpoint** | HTTPS URL where your bot is hosted | `https://mybot.azurewebsites.net` |

The manifest preview will update automatically as you type.

### Step 3: Provisioning

The wizard will automatically:
1. ✅ Create Azure AD application
2. ✅ Generate client secret (2-year expiration)
3. ✅ Create Teams app registration
4. ✅ Register bot with Bot Framework

Progress is shown in real-time with status updates.

### Step 4: Get Credentials

Once complete, you'll receive:

```env
BOT_ID=12345678-1234-1234-1234-123456789abc
CLIENT_ID=12345678-1234-1234-1234-123456789abc
CLIENT_SECRET=abc123...
TENANT_ID=87654321-4321-4321-4321-cba987654321
TEAMS_APP_ID=11111111-2222-3333-4444-555555555555
BOT_ENDPOINT=https://mybot.azurewebsites.net
MESSAGING_ENDPOINT=https://mybot.azurewebsites.net/api/messages
```

**Actions available:**
- Copy individual credentials
- Download `.env` file
- View next steps

## API Endpoints

The backend exposes these REST endpoints:

### Authentication

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/auth/start` | Start device code flow |
| POST | `/api/auth/poll` | Poll for auth completion |
| POST | `/api/auth/token` | Get token for scopes |

### Provisioning

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/provision/aad-app` | Create Azure AD app |
| POST | `/api/provision/client-secret` | Generate client secret |
| POST | `/api/provision/teams-app` | Create Teams app |
| POST | `/api/provision/bot` | Register bot |
| POST | `/api/provision/complete` | Get final credentials |

## Project Structure

```
bot-provisioner-demo/
├── backend/
│   ├── package.json         # Backend dependencies
│   └── server.js            # Express.js server
├── frontend/
│   ├── index.html           # Wizard UI
│   ├── styles.css           # Styling
│   └── app.js               # Frontend logic
└── README.md                # This file
```

## Backend Details

### Technology Stack
- **Express.js** - Web server framework
- **@azure/msal-node** - Microsoft Authentication Library
- **axios** - HTTP client for API calls
- **adm-zip** - Zip file creation for Teams app manifest
- **cors** - Cross-origin resource sharing

### Session Management
Sessions are stored in-memory (Map). For production:
- Use Redis or a database
- Implement session expiration
- Add rate limiting

### Security Considerations

⚠️ **Important:**
- Tokens are stored in-memory (use Redis in production)
- No HTTPS on localhost (use reverse proxy in production)
- CORS is wide open (restrict origins in production)
- No rate limiting (add in production)

## Frontend Details

### Technology Stack
- **Vanilla JavaScript** - No frameworks!
- **CSS Grid/Flexbox** - Responsive layout
- **Fetch API** - HTTP requests

### Features
- Device code authentication flow
- Real-time provisioning progress
- Form validation
- Clipboard API for copying
- File download for .env

### Browser Support
- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+

## Troubleshooting

### "Insufficient privileges to complete the operation"

**Cause:** User doesn't have admin consent for `Application.ReadWrite.All`.

**Solution:**
1. Sign in as Global Administrator
2. Or have admin pre-consent:
   ```bash
   az login
   az ad app permission admin-consent --id <app-id>
   ```

### Backend not starting

**Cause:** Port 3001 already in use.

**Solution:**
```bash
# Find process using port 3001
lsof -ti:3001 | xargs kill -9

# Or change port in server.js:
const PORT = 3002;
```

### Frontend can't reach backend

**Cause:** CORS or network issue.

**Solution:**
1. Check backend is running: `curl http://localhost:3001/health`
2. Check CORS is enabled in `server.js`
3. Open browser DevTools → Network tab for errors

### "Bot already exists" error

**Cause:** Bot ID (AAD app ID) already registered.

**Solution:**
- The wizard will automatically try to update the existing bot
- If that fails, use a different bot name to generate new AAD app

## Development

### Running in Development Mode

```bash
# Backend with auto-reload
cd backend
npm install -g nodemon
npm run dev

# Frontend - just open index.html or use live-server
cd frontend
npx live-server
```

### Adding Features

**Backend:** Add new routes in `server.js`
```javascript
app.post('/api/custom-endpoint', async (req, res) => {
  // Your logic here
});
```

**Frontend:** Add UI in `index.html`, styles in `styles.css`, logic in `app.js`

### Debugging

Enable verbose logging:

```javascript
// In backend/server.js
console.log('Debug:', JSON.stringify(data, null, 2));

// In frontend/app.js
console.log('State:', state);
```

## Production Deployment

### Backend Deployment

1. **Environment Variables**
   ```bash
   export NODE_ENV=production
   export PORT=443
   export SESSION_SECRET=<random-secret>
   ```

2. **HTTPS Setup**
   - Use nginx or Apache as reverse proxy
   - Obtain SSL certificate (Let's Encrypt)

3. **Session Storage**
   - Replace in-memory Map with Redis
   - Add session expiration

4. **Security**
   - Implement rate limiting
   - Add CSRF protection
   - Restrict CORS origins
   - Add request validation

### Frontend Deployment

1. **Static Hosting**
   - Upload to: Azure Static Web Apps, Netlify, Vercel
   - Or serve from backend: `app.use(express.static('frontend'))`

2. **Configuration**
   - Update `API_BASE` in `app.js` to production backend URL
   - Minify JavaScript and CSS

3. **CDN**
   - Optionally add CDN for faster delivery

## Example Use Cases

### CI/CD Pipeline

Use the backend API in your deployment pipeline:

```bash
# Authenticate
SESSION_ID=$(curl -s -X POST http://localhost:3001/api/auth/start \
  -H "Content-Type: application/json" \
  -d '{"scopes": [...]}' | jq -r '.sessionId')

# Provision bot
curl -X POST http://localhost:3001/api/provision/aad-app \
  -H "Content-Type: application/json" \
  -d '{"sessionId": "'$SESSION_ID'", "appName": "CI Bot"}'
```

### Custom Integration

Import backend functions in your Node.js app:

```javascript
const { createAadApp, registerBot } = require('./backend/server.js');
// Use functions programmatically
```

### Bulk Provisioning

Provision multiple bots:

```javascript
const bots = ['Bot1', 'Bot2', 'Bot3'];
for (const botName of bots) {
  await provisionBot(botName, 'https://example.com');
}
```

## Comparison with Teams Toolkit

| Feature | This Demo | Teams Toolkit CLI |
|---------|-----------|-------------------|
| **Installation** | `npm install` (3 packages) | Full toolkit install |
| **UI** | Web wizard | CLI only |
| **Customization** | Full source access | Config files |
| **Integration** | REST API | CLI commands |
| **Learning Curve** | Low | Medium |
| **Production Ready** | Needs hardening | Yes |

## Contributing

This is a demo application. Feel free to:
- Fork and modify
- Add new features
- Improve security
- Add tests
- Submit issues

## License

MIT

## Credits

Built to demonstrate the provisioning flow of the [Microsoft 365 Agents Toolkit](https://github.com/microsoft/teams-toolkit).

## Next Steps

After provisioning your bot:

1. **Develop Bot Logic**
   - Use Bot Framework SDK
   - Implement message handlers
   - Add business logic

2. **Deploy Bot Code**
   - Deploy to Azure App Service
   - Or use Azure Functions
   - Ensure endpoint matches what you configured

3. **Test in Teams**
   - Sideload your app
   - Test bot commands
   - Debug with Bot Framework Emulator

4. **Publish to Teams Store**
   - Prepare app package
   - Submit for review
   - Distribute to users

## Support

For issues or questions:
- Check the Troubleshooting section
- Review backend logs: `DEBUG=* npm start`
- Check browser DevTools console
- Review Microsoft Graph API docs

## Resources

- [Microsoft Graph API](https://docs.microsoft.com/graph/api/overview)
- [Teams Developer Portal API](https://docs.microsoft.com/microsoftteams/platform/concepts/build-and-test/teams-developer-portal)
- [Bot Framework](https://dev.botframework.com/)
- [MSAL Node.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)
