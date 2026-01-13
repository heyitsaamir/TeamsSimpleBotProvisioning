const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const msal = require('@azure/msal-node');
const axios = require('axios');
const AdmZip = require('adm-zip');
const crypto = require('crypto');

const app = express();
const PORT = 3002; // Different port to avoid conflict with public client server

// Middleware
app.use(cors({
    origin: [
        'http://localhost:8080',
        'https://3hvfdfhp-8080.usw2.devtunnels.ms',
        'https://3hvfdfhp-3002.usw2.devtunnels.ms' // Backend tunnel can also receive requests
    ],
    credentials: true
}));
app.use(bodyParser.json());

// Configuration
const CONFIG = {
    clientId: '2a098349-9ecc-463f-a053-d5675e10deeb',
    clientSecret: process.env.CLIENT_SECRET || 'YOUR_CLIENT_SECRET_HERE', // Set via environment variable
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'https://3hvfdfhp-8080.usw2.devtunnels.ms/redirect.html', // Frontend redirect page
    graphBaseUrl: 'https://graph.microsoft.com/v1.0',
    tdpBaseUrl: 'https://dev.teams.microsoft.com',
    graphScopes: [
        'https://graph.microsoft.com/Application.ReadWrite.All',
        'https://graph.microsoft.com/TeamsAppInstallation.ReadForUser'
    ],
    tdpScopes: [
        'https://dev.teams.microsoft.com/AppDefinitions.ReadWrite'
    ],
};

// IMPORTANT: Azure AD doesn't allow requesting scopes from multiple resources in one request
// We can only request Graph scopes OR TDP scopes, not both
// Solution: Request Graph scopes first, then use acquireTokenSilent for TDP (may require admin pre-consent)
const INITIAL_SCOPES = CONFIG.graphScopes; // Start with Graph API scopes only

// In-memory session storage (use Redis in production)
const sessions = new Map();

// Store PKCE code verifiers and states
const pendingAuthRequests = new Map();

// MSAL Confidential Client
const msalConfig = {
    auth: {
        clientId: CONFIG.clientId,
        authority: CONFIG.authority,
        clientSecret: CONFIG.clientSecret,
    }
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

// ═══════════════════════════════════════════════════════════════
// AUTHENTICATION ENDPOINTS
// ═══════════════════════════════════════════════════════════════

/**
 * GET /api/auth/start
 * Start authorization code flow (confidential client)
 * Returns authorization URL for user to visit
 */
app.get('/api/auth/start', async (req, res) => {
    try {
        // Generate PKCE code verifier and challenge
        const codeVerifier = generateCodeVerifier();
        const codeChallenge = generateCodeChallenge(codeVerifier);

        // Generate state for CSRF protection
        const state = generateSessionId();

        // Store PKCE verifier and state for later verification
        pendingAuthRequests.set(state, {
            codeVerifier,
            createdAt: Date.now(),
        });

        // Build authorization URL
        // IMPORTANT: Can only request scopes from ONE resource per authorization
        // We request Graph scopes here, TDP will need separate consent
        const authCodeUrlParameters = {
            scopes: INITIAL_SCOPES, // Graph scopes only (can't mix resources)
            redirectUri: CONFIG.redirectUri,
            codeChallenge: codeChallenge,
            codeChallengeMethod: 'S256',
            state: state,
        };

        const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);

        console.log('Generated auth URL with state:', state);

        res.json({
            authUrl: authUrl,
            state: state,
        });

    } catch (error) {
        console.error('Auth start error:', error);
        res.status(500).json({ error: error.message });
    }
});

/**
 * POST /api/auth/callback
 * Handle OAuth callback and exchange code for tokens
 * Called by frontend after redirect
 */
app.post('/api/auth/callback', async (req, res) => {
    const { code, state } = req.body;

    try {
        // Verify state to prevent CSRF
        const pendingRequest = pendingAuthRequests.get(state);
        if (!pendingRequest) {
            return res.status(400).json({ error: 'Invalid or expired state parameter' });
        }

        // Exchange authorization code for tokens
        // Use same scopes as authorization request (Graph only)
        const tokenRequest = {
            code: code,
            scopes: INITIAL_SCOPES, // Graph scopes only
            redirectUri: CONFIG.redirectUri,
            codeVerifier: pendingRequest.codeVerifier,
        };

        console.log('Exchanging authorization code for tokens...');

        const response = await cca.acquireTokenByCode(tokenRequest);

        // Get the account from the token response
        const account = response.account;

        // Create new session with account info
        const sessionId = generateSessionId();
        sessions.set(sessionId, {
            account: account,
            createdAt: Date.now(),
        });

        // Clean up the pending request
        pendingAuthRequests.delete(state);

        console.log('Created session:', sessionId);
        console.log('Authenticated as:', account.username);
        console.log('Tokens cached by MSAL for acquireTokenSilent');

        res.json({
            success: true,
            sessionId: sessionId,
            userInfo: {
                username: account.username,
                tenantId: account.tenantId,
            }
        });

    } catch (error) {
        console.error('Auth callback error:', error);
        res.status(500).json({ error: error.message });
    }
});

/**
 * POST /api/auth/token
 * Get token for specific scopes using acquireTokenSilent()
 */
app.post('/api/auth/token', async (req, res) => {
    const { sessionId, scopes } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    if (!session.account) {
        return res.status(401).json({ error: 'No account in session' });
    }

    try {
        // Use acquireTokenSilent to get token for specific scopes
        const silentRequest = {
            account: session.account,
            scopes: scopes,
            forceRefresh: false,
        };

        console.log('Acquiring token silently for scopes:', scopes.join(', '));

        const response = await cca.acquireTokenSilent(silentRequest);

        console.log('Token acquired successfully for:', session.account.username);

        res.json({ token: response.accessToken });

    } catch (error) {
        console.error('acquireTokenSilent error:', error);
        res.status(403).json({
            error: 'Failed to acquire token silently',
            details: error.message,
            needsAuth: true
        });
    }
});

// ═══════════════════════════════════════════════════════════════
// PROVISIONING ENDPOINTS (same as public client)
// ═══════════════════════════════════════════════════════════════

/**
 * POST /api/provision/aad-app
 * Create Azure AD application using acquireTokenSilent
 */
app.post('/api/provision/aad-app', async (req, res) => {
    const { sessionId, appName } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        // Get token for Graph API scopes using acquireTokenSilent
        const token = await getTokenForScopes(sessionId, CONFIG.graphScopes);

        // Create AAD App
        const response = await axios.post(
            `${CONFIG.graphBaseUrl}/applications`,
            {
                displayName: appName,
                signInAudience: 'AzureADMultipleOrgs',
            },
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                }
            }
        );

        const app = response.data;

        res.json({
            clientId: app.appId,
            objectId: app.id,
        });

    } catch (error) {
        console.error('AAD app creation error:', error.response?.data || error.message);
        res.status(500).json({ error: error.response?.data || error.message });
    }
});

/**
 * POST /api/provision/client-secret
 * Generate client secret for AAD app using acquireTokenSilent
 */
app.post('/api/provision/client-secret', async (req, res) => {
    const { sessionId, objectId } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        // Get token for Graph API scopes using acquireTokenSilent
        const token = await getTokenForScopes(sessionId, CONFIG.graphScopes);

        const expireDate = new Date();
        expireDate.setFullYear(expireDate.getFullYear() + 2);

        const response = await axios.post(
            `${CONFIG.graphBaseUrl}/applications/${objectId}/addPassword`,
            {
                passwordCredential: {
                    displayName: 'default',
                    endDateTime: expireDate.toISOString(),
                }
            },
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                }
            }
        );

        res.json({
            clientSecret: response.data.secretText,
            expiresOn: expireDate.toISOString(),
        });

    } catch (error) {
        console.error('Secret generation error:', error.response?.data || error.message);
        res.status(500).json({ error: error.response?.data || error.message });
    }
});

/**
 * POST /api/provision/teams-app
 * Create Teams app using acquireTokenSilent
 */
app.post('/api/provision/teams-app', async (req, res) => {
    const { sessionId, manifest } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        // Get token for Teams Dev Portal scopes using acquireTokenSilent
        const token = await getTokenForScopes(sessionId, CONFIG.tdpScopes);

        // Create zip file with manifest
        const zip = new AdmZip();
        zip.addFile('manifest.json', Buffer.from(JSON.stringify(manifest, null, 2)));

        // Add placeholder icons
        const colorIcon = createPlaceholderPng();
        const outlineIcon = createPlaceholderPng();
        zip.addFile('color.png', colorIcon);
        zip.addFile('outline.png', outlineIcon);

        const zipBuffer = zip.toBuffer();

        // Upload to Teams Dev Portal
        const response = await axios.post(
            `${CONFIG.tdpBaseUrl}/api/appdefinitions/v2/import`,
            zipBuffer,
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/zip',
                    'Client-Source': 'teamstoolkit',
                }
            }
        );

        res.json({
            teamsAppId: response.data.teamsAppId,
            tenantId: response.data.tenantId,
        });

    } catch (error) {
        console.error('Teams app creation error:', error.response?.data || error.message);
        res.status(500).json({ error: error.response?.data || error.message });
    }
});

/**
 * POST /api/provision/bot
 * Register bot with Bot Framework using acquireTokenSilent
 */
app.post('/api/provision/bot', async (req, res) => {
    const { sessionId, botId, botName, messagingEndpoint } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        // Get token for Teams Dev Portal scopes using acquireTokenSilent
        const token = await getTokenForScopes(sessionId, CONFIG.tdpScopes);

        const response = await axios.post(
            `${CONFIG.tdpBaseUrl}/api/botframework`,
            {
                botId: botId,
                name: botName,
                description: '',
                messagingEndpoint: messagingEndpoint,
                callingEndpoint: '',
                configuredChannels: ['msteams'],
                isSingleTenant: true,
            },
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                    'Client-Source': 'teamstoolkit',
                }
            }
        );

        res.json({ success: true });

    } catch (error) {
        // Bot might already exist, try update
        if (error.response?.status === 409) {
            try {
                const token = await getTokenForScopes(sessionId, CONFIG.tdpScopes);
                await axios.post(
                    `${CONFIG.tdpBaseUrl}/api/botframework/${botId}`,
                    {
                        botId: botId,
                        name: botName,
                        messagingEndpoint: messagingEndpoint,
                        configuredChannels: ['msteams'],
                    },
                    {
                        headers: {
                            'Authorization': `Bearer ${token}`,
                            'Content-Type': 'application/json',
                            'Client-Source': 'teamstoolkit',
                        }
                    }
                );
                res.json({ success: true, updated: true });
            } catch (updateError) {
                console.error('Bot update error:', updateError.response?.data || updateError.message);
                res.status(500).json({ error: updateError.response?.data || updateError.message });
            }
        } else {
            console.error('Bot registration error:', error.response?.data || error.message);
            res.status(500).json({ error: error.response?.data || error.message });
        }
    }
});

/**
 * POST /api/provision/complete
 * Complete provisioning and return all credentials
 */
app.post('/api/provision/complete', async (req, res) => {
    const { sessionId, credentials } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    // Get tenant ID from cached account
    const tenantId = session.account ? session.account.tenantId : null;

    const fullCredentials = {
        ...credentials,
        TENANT_ID: tenantId,
    };

    res.json({
        success: true,
        credentials: fullCredentials,
    });
});

// ═══════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Get token for specific scopes using acquireTokenSilent
 * NOTE: For TDP scopes, admin must pre-consent since we can't request multiple resources in initial auth
 */
async function getTokenForScopes(sessionId, scopes) {
    const session = sessions.get(sessionId);
    if (!session || !session.account) {
        throw new Error('Invalid session or no account');
    }

    try {
        const silentRequest = {
            account: session.account,
            scopes: scopes,
            forceRefresh: false,
        };

        console.log(`Getting token for scopes: ${scopes.join(', ')}`);

        const response = await cca.acquireTokenSilent(silentRequest);
        return response.accessToken;

    } catch (error) {
        console.error('Failed to acquire token silently:', error.message);

        // If TDP scopes fail, it's likely because admin hasn't pre-consented
        const isTdpScope = scopes.some(s => s.includes('dev.teams.microsoft.com'));
        if (isTdpScope && error.message.includes('AADSTS65001')) {
            throw new Error('Teams Dev Portal consent required. Admin must pre-consent to AppDefinitions.ReadWrite scope in Azure AD. See README for instructions.');
        }

        throw error;
    }
}

function generateSessionId() {
    return Math.random().toString(36).substring(2) + Date.now().toString(36);
}

/**
 * Generate PKCE code verifier (random string)
 */
function generateCodeVerifier() {
    return crypto.randomBytes(32).toString('base64url');
}

/**
 * Generate PKCE code challenge (SHA256 hash of verifier)
 */
function generateCodeChallenge(verifier) {
    return crypto
        .createHash('sha256')
        .update(verifier)
        .digest('base64url');
}

function createPlaceholderPng() {
    return Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
        'base64'
    );
}

// Health check
app.get('/health', (req, res) => {
    res.json({ status: 'ok', sessions: sessions.size });
});

// Start server
app.listen(PORT, () => {
    console.log(`\n🚀 Bot Provisioner Backend (CONFIDENTIAL CLIENT) running on http://localhost:${PORT}`);
    console.log(`\n⚙️  Configuration:`);
    console.log(`   Client ID: ${CONFIG.clientId}`);
    console.log(`   Redirect URI: ${CONFIG.redirectUri}`);
    console.log(`   Client Secret: ${CONFIG.clientSecret ? '✓ Set' : '✗ NOT SET - Check environment variable'}`);
    console.log(`\n📋 Available endpoints:`);
    console.log(`   GET  /api/auth/start            - Get authorization URL`);
    console.log(`   POST /api/auth/callback         - Handle OAuth callback`);
    console.log(`   POST /api/provision/aad-app     - Create AAD app`);
    console.log(`   POST /api/provision/client-secret - Generate secret`);
    console.log(`   POST /api/provision/teams-app   - Create Teams app`);
    console.log(`   POST /api/provision/bot         - Register bot`);
    console.log(`   POST /api/provision/complete    - Get final credentials\n`);

    if (!CONFIG.clientSecret || CONFIG.clientSecret === 'YOUR_CLIENT_SECRET_HERE') {
        console.log(`⚠️  WARNING: Client secret not configured!`);
        console.log(`   Set CLIENT_SECRET environment variable before running.`);
        console.log(`   Example: CLIENT_SECRET=your-secret node server-confidential.js\n`);
    }

    console.log(`⚠️  IMPORTANT: Teams Dev Portal Consent`);
    console.log(`   Azure AD doesn't allow requesting scopes from multiple resources in one auth flow.`);
    console.log(`   Initial auth requests Graph API scopes only.`);
    console.log(`   For Teams Dev Portal (AppDefinitions.ReadWrite), you must:`);
    console.log(`   1. Go to Azure AD → App registrations → ${CONFIG.clientId}`);
    console.log(`   2. API permissions → Add permission → APIs my org uses → "Microsoft Teams"`);
    console.log(`   3. Search for "AppDefinitions.ReadWrite" → Add permission`);
    console.log(`   4. Click "Grant admin consent" button`);
    console.log(`   Without admin pre-consent, bot provisioning will fail at Teams app creation.\n`);
});
