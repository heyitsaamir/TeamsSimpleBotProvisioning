const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const msal = require('@azure/msal-node');
const axios = require('axios');
const AdmZip = require('adm-zip');

const app = express();
const PORT = 3003;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Configuration
const ENDPOINT_BASE = process.env.ENDPOINT_BASE || 'https://3hvfdfhp-8080.usw2.devtunnels.ms';

const CONFIG = {
    clientId: '22eb633f-f8bb-4818-90a6-ffac6de52b01',
    clientSecret: process.env.CLIENT_SECRET || 'YOUR_CLIENT_SECRET',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: `${ENDPOINT_BASE}/redirect.html`,
    adminConsentRedirectUri: `${ENDPOINT_BASE}/admin-consent-callback.html`,
    graphBaseUrl: 'https://graph.microsoft.com/v1.0',
    tdpBaseUrl: 'https://dev.teams.microsoft.com',
    graphScopes: [
        'https://graph.microsoft.com/Application.ReadWrite.All'
    ],
    tdpScopes: [
        'https://dev.teams.microsoft.com/AppDefinitions.ReadWrite'
    ],
};

// In-memory session storage
const sessions = new Map();

// Confidential Client Application
const cca = new msal.ConfidentialClientApplication({
    auth: {
        clientId: CONFIG.clientId,
        authority: CONFIG.authority,
        clientSecret: CONFIG.clientSecret,
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Info,
        }
    }
});

// ═══════════════════════════════════════════════════════════════
// AUTHENTICATION ENDPOINTS (CONFIDENTIAL CLIENT)
// ═══════════════════════════════════════════════════════════════

/**
 * GET /api/auth/start
 * Start OAuth authorization code flow with User.Read scope only
 */
app.get('/api/auth/start', async (req, res) => {
    try {
        const state = generateSessionId();

        const authCodeUrlParameters = {
            scopes: ['User.Read'], // Minimal scope - always works
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
 * Handle OAuth callback and exchange code for tokens
 */
app.post('/api/auth/callback', async (req, res) => {
    const { code, state } = req.body;

    try {
        const tokenRequest = {
            code: code,
            scopes: ['User.Read'],
            redirectUri: CONFIG.redirectUri,
        };

        const response = await cca.acquireTokenByCode(tokenRequest);

        // Create session
        const sessionId = generateSessionId();
        sessions.set(sessionId, {
            account: response.account,
            createdAt: Date.now(),
        });

        console.log('User authenticated:', response.account.username);

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

/**
 * POST /api/auth/check-consent
 * Check if admin has granted all required permissions by trying to acquire tokens
 */
app.post('/api/auth/check-consent', async (req, res) => {
    const { sessionId } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        // Try to acquire token silently for Graph scopes
        // If admin consent is granted, this will succeed
        // If not, it will fail with consent error

        const grantedScopes = [];
        const missingScopes = [];
        const scopeErrors = {}; // Track error reasons for debugging

        // Test Graph scopes
        for (const scope of CONFIG.graphScopes) {
            try {
                const silentRequest = {
                    account: session.account,
                    scopes: [scope],
                    forceRefresh: false,
                };
                await cca.acquireTokenSilent(silentRequest);
                grantedScopes.push(scope.split('/').pop()); // Extract scope name
                console.log(`✓ Scope granted: ${scope}`);
            } catch (error) {
                const scopeName = scope.split('/').pop();
                const errorCode = error.errorCode || error.name;
                const errorMessage = error.message || '';

                // Check if it's an expected consent error vs unexpected error
                const isConsentError =
                    errorCode === 'consent_required' ||
                    errorCode === 'interaction_required' ||
                    (errorCode === 'invalid_grant' && (errorMessage.includes('AADSTS65001') || errorMessage.includes('has not consented')));

                if (isConsentError) {
                    // Expected: Admin hasn't consented yet
                    missingScopes.push(scopeName);
                    scopeErrors[scopeName] = errorCode;
                    console.log(`✗ Scope missing (needs consent): ${scope} - ${errorCode}`);
                } else {
                    // Unexpected error - log and fail
                    console.error(`❌ Unexpected error for scope ${scope}:`, errorCode, error.message);
                    throw new Error(`Failed to check scope ${scopeName}: ${errorCode || error.message}`);
                }
            }
        }

        // Test TDP scopes
        for (const scope of CONFIG.tdpScopes) {
            try {
                const silentRequest = {
                    account: session.account,
                    scopes: [scope],
                    forceRefresh: false,
                };
                await cca.acquireTokenSilent(silentRequest);
                grantedScopes.push(scope.split('/').pop());
                console.log(`✓ Scope granted: ${scope}`);
            } catch (error) {
                const scopeName = scope.split('/').pop();
                const errorCode = error.errorCode || error.name;
                const errorMessage = error.message || '';

                // Check if it's an expected consent error vs unexpected error
                const isConsentError =
                    errorCode === 'consent_required' ||
                    errorCode === 'interaction_required' ||
                    (errorCode === 'invalid_grant' && (errorMessage.includes('AADSTS65001') || errorMessage.includes('has not consented')));

                if (isConsentError) {
                    // Expected: Admin hasn't consented yet
                    missingScopes.push(scopeName);
                    scopeErrors[scopeName] = errorCode;
                    console.log(`✗ Scope missing (needs consent): ${scope} - ${errorCode}`);
                } else {
                    // Unexpected error - log and fail
                    console.error(`❌ Unexpected error for scope ${scope}:`, errorCode, error.message);
                    throw new Error(`Failed to check scope ${scopeName}: ${errorCode || error.message}`);
                }
            }
        }

        if (missingScopes.length > 0) {
            return res.json({
                hasConsent: false,
                missingScopes: missingScopes,
                grantedScopes: grantedScopes,
                scopeErrors: scopeErrors, // Include error details for debugging
                adminConsentUrl: `https://login.microsoftonline.com/${session.account.tenantId}/adminconsent?client_id=${CONFIG.clientId}&redirect_uri=${encodeURIComponent(CONFIG.adminConsentRedirectUri)}`,
            });
        }

        console.log('✓ All required scopes granted');
        res.json({
            hasConsent: true,
            grantedScopes: grantedScopes,
        });

    } catch (error) {
        console.error('Check consent error:', error.response?.data || error.message);
        res.status(500).json({ error: error.message });
    }
});

/**
 * POST /api/check-sideloading
 * Check if tenant has sideloading enabled using Teams Dev Portal API
 */
app.post('/api/check-sideloading', async (req, res) => {
    const { sessionId } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
        // Get token for Teams Dev Portal
        const token = await getTokenForScopes(sessionId, CONFIG.tdpScopes);

        // Call Teams Dev Portal API to check sideloading status
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

        // If it's a consent error, tell user they need TDP access
        if (error.message?.includes('consent_required') || error.message?.includes('interaction_required')) {
            return res.status(403).json({
                error: 'Teams Dev Portal access required',
                needsConsent: true
            });
        }

        res.status(500).json({ error: error.message });
    }
});

// ═══════════════════════════════════════════════════════════════
// PROVISIONING ENDPOINTS (REUSED FROM server.js)
// ═══════════════════════════════════════════════════════════════

/**
 * POST /api/provision/aad-app
 * Create Azure AD application
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
 * Generate client secret for AAD app
 */
app.post('/api/provision/client-secret', async (req, res) => {
    const { sessionId, objectId } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
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
 * Create Teams app
 */
app.post('/api/provision/teams-app', async (req, res) => {
    const { sessionId, manifest } = req.body;

    const session = sessions.get(sessionId);
    if (!session) {
        return res.status(401).json({ error: 'Invalid session' });
    }

    try {
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
 * Register bot with Bot Framework
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
        throw error;
    }
}

function generateSessionId() {
    return Math.random().toString(36).substring(2) + Date.now().toString(36);
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
    console.log(`\n🚀 Bot Provisioner Backend (Confidential Client) running on http://localhost:${PORT}`);
    console.log(`\n📋 Available endpoints:`);
    console.log(`   GET  /api/auth/start           - Start OAuth flow (User.Read)`);
    console.log(`   POST /api/auth/callback        - OAuth callback handler`);
    console.log(`   POST /api/auth/check-consent   - Check admin consent status`);
    console.log(`   POST /api/check-sideloading    - Check tenant sideloading status`);
    console.log(`   POST /api/provision/aad-app    - Create AAD app`);
    console.log(`   POST /api/provision/client-secret - Generate secret`);
    console.log(`   POST /api/provision/teams-app  - Create Teams app`);
    console.log(`   POST /api/provision/bot        - Register bot`);
    console.log(`   POST /api/provision/complete   - Get final credentials\n`);
});
