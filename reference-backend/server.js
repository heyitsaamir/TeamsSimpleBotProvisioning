/**
 * Bot Provisioner - Reference Backend Implementation
 *
 * This server demonstrates how to implement a confidential OAuth client that:
 * 1. Authenticates users via OAuth 2.0 authorization code flow
 * 2. Checks if admin consent has been granted for required scopes
 * 3. Provisions Teams bots by calling Microsoft Graph and Teams Developer Portal APIs
 *
 * Key Concepts:
 * - Confidential Client: Server-side app that can securely store a client secret
 * - Authorization Code Flow: Standard OAuth flow for web applications
 * - Multi-tenant: App is registered in one tenant but can be used by any tenant
 * - Silent Token Acquisition: Use refresh tokens to get access tokens without user interaction
 */

const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const msal = require('@azure/msal-node');
const axios = require('axios');
const AdmZip = require('adm-zip');

const app = express();
const PORT = process.env.PORT || 3003;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// ═══════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════

/**
 * Application Configuration
 *
 * These values come from your Azure AD app registration.
 * In production, use environment variables for sensitive data.
 */
const CONFIG = {
    // Your multi-tenant app's client ID from Azure Portal
    clientId: process.env.CLIENT_ID || 'YOUR_CLIENT_ID',

    // Client secret generated in Azure Portal (keep this secure!)
    clientSecret: process.env.CLIENT_SECRET || 'YOUR_CLIENT_SECRET',

    // Authority: /common allows any Azure AD tenant to authenticate
    authority: 'https://login.microsoftonline.com/common',

    // Where Azure AD redirects after user authenticates
    redirectUri: process.env.REDIRECT_URI || 'http://localhost:8080/redirect.html',

    // Where Azure AD redirects after admin grants consent
    adminConsentRedirectUri: process.env.ADMIN_CONSENT_URI || 'http://localhost:8080/admin-consent-callback.html',

    // API endpoints
    graphBaseUrl: 'https://graph.microsoft.com/v1.0',
    tdpBaseUrl: 'https://dev.teams.microsoft.com',

    // Required scopes for Microsoft Graph API
    graphScopes: ['https://graph.microsoft.com/Application.ReadWrite.All'],

    // Required scopes for Teams Developer Portal API
    tdpScopes: ['https://dev.teams.microsoft.com/AppDefinitions.ReadWrite'],
};

// ═══════════════════════════════════════════════════════════════
// MSAL SETUP
// ═══════════════════════════════════════════════════════════════

/**
 * MSAL Confidential Client Application
 *
 * This is the core authentication component. MSAL handles:
 * - Generating OAuth authorization URLs
 * - Exchanging authorization codes for tokens
 * - Token caching and refresh token management
 * - Silent token acquisition using cached refresh tokens
 */
const cca = new msal.ConfidentialClientApplication({
    auth: {
        clientId: CONFIG.clientId,
        authority: CONFIG.authority,
        clientSecret: CONFIG.clientSecret,
    },
});

/**
 * In-Memory Session Storage
 *
 * Stores user account information keyed by sessionId.
 * In production, use a persistent store like Redis.
 *
 * Structure: Map<sessionId, { account: MSALAccount, createdAt: timestamp }>
 */
const sessions = new Map();

// ═══════════════════════════════════════════════════════════════
// AUTHENTICATION ENDPOINTS
// ═══════════════════════════════════════════════════════════════

/**
 * GET /api/auth/start
 *
 * Initiates the OAuth authorization code flow.
 *
 * Flow:
 * 1. Generate a unique state parameter for CSRF protection
 * 2. Create authorization URL with User.Read scope (minimal, always grantable)
 * 3. Return URL to frontend, which redirects user to Azure AD
 *
 * Why User.Read?
 * - It's a basic scope that any user can consent to
 * - Allows users to sign in even if admin consent is missing
 * - We'll check for admin-consented scopes separately
 */
app.get('/api/auth/start', async (req, res) => {
    try {
        const state = generateSessionId();

        const authCodeUrlParameters = {
            scopes: ['User.Read'],
            redirectUri: CONFIG.redirectUri,
            state: state,
        };

        const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);

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
 *
 * Handles the OAuth callback after user authenticates.
 *
 * Flow:
 * 1. Receive authorization code from Azure AD
 * 2. Exchange code for access token + refresh token
 * 3. Extract account information from response
 * 4. Store account in session (enables silent token acquisition later)
 * 5. Return sessionId to frontend
 *
 * The account object contains:
 * - homeAccountId: Unique identifier for the user
 * - username: User's email/UPN
 * - tenantId: User's tenant ID
 *
 * MSAL automatically caches the refresh token internally.
 */
app.post('/api/auth/callback', async (req, res) => {
    const { code, state } = req.body;

    if (!code) {
        return res.status(400).json({ error: 'Missing authorization code' });
    }

    try {
        const tokenRequest = {
            code: code,
            scopes: ['User.Read'],
            redirectUri: CONFIG.redirectUri,
        };

        // Exchange authorization code for tokens
        const response = await cca.acquireTokenByCode(tokenRequest);

        // Create session with account information
        const sessionId = generateSessionId();
        sessions.set(sessionId, {
            account: response.account,
            createdAt: Date.now(),
        });

        console.log(`User authenticated: ${response.account.username}`);

        res.json({
            sessionId: sessionId,
            userInfo: {
                username: response.account.username,
                tenantId: response.account.tenantId,
            }
        });

    } catch (error) {
        console.error('Auth callback error:', error);
        res.status(500).json({ error: error.message });
    }
});

// ═══════════════════════════════════════════════════════════════
// SCOPE CHECKING
// ═══════════════════════════════════════════════════════════════

/**
 * POST /api/auth/check-consent
 *
 * Checks if admin has granted consent for required scopes.
 *
 * How it works:
 * 1. For each required scope, attempt acquireTokenSilent()
 * 2. If successful → scope is granted
 * 3. If fails with consent error → scope needs admin consent
 * 4. If fails with other error → unexpected problem
 *
 * Why acquireTokenSilent?
 * - Uses cached refresh token to request new access tokens
 * - Doesn't require user interaction
 * - Fails predictably when consent is missing
 * - Works across different resource servers (Graph, TDP)
 *
 * Expected consent error codes:
 * - consent_required: Scope needs consent
 * - interaction_required: User interaction needed (usually consent)
 * - invalid_grant with AADSTS65001: Admin hasn't consented
 */
app.post('/api/auth/check-consent', async (req, res) => {
    const { sessionId } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        const grantedScopes = [];
        const missingScopes = [];
        const scopeErrors = {};

        // Check Microsoft Graph scopes
        for (const scope of CONFIG.graphScopes) {
            try {
                await cca.acquireTokenSilent({
                    account: session.account,
                    scopes: [scope],
                    forceRefresh: false,
                });

                const scopeName = scope.split('/').pop();
                grantedScopes.push(scopeName);
                console.log(`✓ Scope granted: ${scope}`);

            } catch (error) {
                const scopeName = scope.split('/').pop();
                const errorCode = error.errorCode || error.name;
                const errorMessage = error.message || '';

                // Is this an expected consent error?
                const isConsentError =
                    errorCode === 'consent_required' ||
                    errorCode === 'interaction_required' ||
                    (errorCode === 'invalid_grant' &&
                     (errorMessage.includes('AADSTS65001') || errorMessage.includes('has not consented')));

                if (isConsentError) {
                    missingScopes.push(scopeName);
                    scopeErrors[scopeName] = errorCode;
                    console.log(`✗ Scope missing: ${scope}`);
                } else {
                    // Unexpected error - fail fast
                    console.error(`❌ Unexpected error for scope ${scope}:`, errorCode, errorMessage);
                    throw new Error(`Failed to check scope ${scopeName}: ${errorCode || error.message}`);
                }
            }
        }

        // Check Teams Developer Portal scopes
        for (const scope of CONFIG.tdpScopes) {
            try {
                await cca.acquireTokenSilent({
                    account: session.account,
                    scopes: [scope],
                    forceRefresh: false,
                });

                const scopeName = scope.split('/').pop();
                grantedScopes.push(scopeName);
                console.log(`✓ Scope granted: ${scope}`);

            } catch (error) {
                const scopeName = scope.split('/').pop();
                const errorCode = error.errorCode || error.name;
                const errorMessage = error.message || '';

                const isConsentError =
                    errorCode === 'consent_required' ||
                    errorCode === 'interaction_required' ||
                    (errorCode === 'invalid_grant' &&
                     (errorMessage.includes('AADSTS65001') || errorMessage.includes('has not consented')));

                if (isConsentError) {
                    missingScopes.push(scopeName);
                    scopeErrors[scopeName] = errorCode;
                    console.log(`✗ Scope missing: ${scope}`);
                } else {
                    console.error(`❌ Unexpected error for scope ${scope}:`, errorCode, errorMessage);
                    throw new Error(`Failed to check scope ${scopeName}: ${errorCode || error.message}`);
                }
            }
        }

        // If scopes are missing, generate admin consent URL
        if (missingScopes.length > 0) {
            return res.json({
                hasConsent: false,
                missingScopes: missingScopes,
                grantedScopes: grantedScopes,
                scopeErrors: scopeErrors,
                adminConsentUrl: `https://login.microsoftonline.com/${session.account.tenantId}/adminconsent?client_id=${CONFIG.clientId}&redirect_uri=${encodeURIComponent(CONFIG.adminConsentRedirectUri)}`,
            });
        }

        console.log('✓ All required scopes granted');
        res.json({
            hasConsent: true,
            grantedScopes: grantedScopes,
        });

    } catch (error) {
        console.error('Check consent error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// ═══════════════════════════════════════════════════════════════
// SIDELOADING CHECK
// ═══════════════════════════════════════════════════════════════

/**
 * POST /api/check-sideloading
 *
 * Checks if the user's tenant allows custom app sideloading.
 *
 * This is required for users to install the provisioned bot in Teams.
 * If sideloading is disabled, the tenant admin must enable it in the
 * Teams admin center.
 */
app.post('/api/check-sideloading', async (req, res) => {
    const { sessionId } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        // Get token for Teams Developer Portal
        const token = await getTokenForScopes(sessionId, CONFIG.tdpScopes);

        // Call TDP API to check sideloading policy
        const response = await axios.get(
            `${CONFIG.tdpBaseUrl}/api/usersettings/mtUserAppPolicy`,
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                }
            }
        );

        const isSideloadingAllowed = response.data?.value?.isSideloadingAllowed;

        console.log(`Sideloading status: ${isSideloadingAllowed ? 'Enabled' : 'Disabled'}`);

        res.json({
            isSideloadingAllowed: isSideloadingAllowed,
            status: isSideloadingAllowed === true ? 'enabled' :
                    isSideloadingAllowed === false ? 'disabled' : 'unknown'
        });

    } catch (error) {
        console.error('Check sideloading error:', error.response?.data || error.message);
        res.status(500).json({ error: error.message });
    }
});

// ═══════════════════════════════════════════════════════════════
// PROVISIONING ENDPOINTS
// ═══════════════════════════════════════════════════════════════

/**
 * POST /api/provision/aad-app
 *
 * Creates an Azure AD app registration.
 *
 * This app will serve as the bot's identity. The response includes:
 * - clientId (appId): The application's public identifier
 * - appRegistrationId (id): Internal Azure AD object ID for management operations
 */
app.post('/api/provision/aad-app', async (req, res) => {
    const { sessionId, appName } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        const token = await getTokenForScopes(sessionId, CONFIG.graphScopes);

        const response = await axios.post(
            `${CONFIG.graphBaseUrl}/applications`,
            {
                displayName: appName,
                signInAudience: 'AzureADMultipleOrgs', // Multi-tenant
            },
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                }
            }
        );

        const app = response.data;

        console.log(`Created AAD app: ${app.appId}`);

        res.json({
            clientId: app.appId,
            appRegistrationId: app.id, // Used for subsequent operations
        });

    } catch (error) {
        console.error('AAD app creation error:', error.response?.data || error.message);
        res.status(500).json({ error: error.response?.data || error.message });
    }
});

/**
 * POST /api/provision/client-secret
 *
 * Generates a client secret for the Azure AD app.
 *
 * Important: The secret value is only returned once. Store it securely.
 * The secret expires after 2 years by default (configurable).
 */
app.post('/api/provision/client-secret', async (req, res) => {
    const { sessionId, appRegistrationId } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        const token = await getTokenForScopes(sessionId, CONFIG.graphScopes);

        // Set expiration to 2 years from now
        const expireDate = new Date();
        expireDate.setFullYear(expireDate.getFullYear() + 2);

        const response = await axios.post(
            `${CONFIG.graphBaseUrl}/applications/${appRegistrationId}/addPassword`,
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

        console.log(`Generated client secret for app: ${appRegistrationId}`);

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
 *
 * Creates a Teams app package and uploads to Teams Developer Portal.
 *
 * The package must be a zip file containing:
 * - manifest.json: App definition (bots, tabs, etc.)
 * - color.png: 192x192 color icon
 * - outline.png: 32x32 outline icon
 */
app.post('/api/provision/teams-app', async (req, res) => {
    const { sessionId, manifest } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        const token = await getTokenForScopes(sessionId, CONFIG.tdpScopes);

        // Create zip file with manifest and placeholder icons
        const zip = new AdmZip();
        zip.addFile('manifest.json', Buffer.from(JSON.stringify(manifest, null, 2)));
        zip.addFile('color.png', createPlaceholderPng());
        zip.addFile('outline.png', createPlaceholderPng());

        const zipBuffer = zip.toBuffer();

        // Upload to Teams Dev Portal
        const response = await axios.post(
            `${CONFIG.tdpBaseUrl}/api/appdefinitions/v2/import`,
            zipBuffer,
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/zip',
                }
            }
        );

        console.log(`Created Teams app: ${response.data.teamsAppId}`);

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
 *
 * Registers a bot with the Bot Framework via Teams Developer Portal.
 *
 * The bot must have:
 * - botId: The Azure AD app's client ID
 * - messagingEndpoint: HTTPS URL where bot receives messages (e.g., https://yourapp.com/api/messages)
 * - configuredChannels: ['msteams'] for Teams-only bots
 */
app.post('/api/provision/bot', async (req, res) => {
    const { sessionId, botId, botName, messagingEndpoint } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
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
                }
            }
        );

        console.log(`Registered bot: ${botId}`);

        res.json({ success: true });

    } catch (error) {
        // Bot might already exist (409 conflict)
        if (error.response?.status === 409) {
            try {
                // Update existing bot instead
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
                        }
                    }
                );
                console.log(`Updated existing bot: ${botId}`);
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

// ═══════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Acquires an access token for the specified scopes using silent acquisition.
 *
 * This uses the cached refresh token to get a new access token without
 * requiring user interaction. Works across different resource servers.
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

        const response = await cca.acquireTokenSilent(silentRequest);
        return response.accessToken;

    } catch (error) {
        console.error('Failed to acquire token silently:', error.message);
        throw error;
    }
}

/**
 * Generates a random session ID for CSRF protection and session management.
 */
function generateSessionId() {
    return Math.random().toString(36).substring(2) + Date.now().toString(36);
}

/**
 * Creates a minimal 1x1 PNG image for placeholder icons.
 * In production, use proper 192x192 and 32x32 icons.
 */
function createPlaceholderPng() {
    return Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
        'base64'
    );
}

// ═══════════════════════════════════════════════════════════════
// SERVER STARTUP
// ═══════════════════════════════════════════════════════════════

app.listen(PORT, () => {
    console.log(`\n🚀 Bot Provisioner Backend running on http://localhost:${PORT}`);
    console.log(`\n📋 Available endpoints:`);
    console.log(`   GET  /api/auth/start           - Start OAuth flow`);
    console.log(`   POST /api/auth/callback        - OAuth callback handler`);
    console.log(`   POST /api/auth/check-consent   - Check admin consent status`);
    console.log(`   POST /api/check-sideloading    - Check tenant sideloading`);
    console.log(`   POST /api/provision/aad-app    - Create AAD app`);
    console.log(`   POST /api/provision/client-secret - Generate secret`);
    console.log(`   POST /api/provision/teams-app  - Create Teams app`);
    console.log(`   POST /api/provision/bot        - Register bot\n`);
});
