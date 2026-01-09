const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const msal = require('@azure/msal-node');
const axios = require('axios');
const AdmZip = require('adm-zip');

const app = express();
const PORT = 3001;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Configuration
const CONFIG = {
  clientId: '7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0', // Actual CLI client ID (same as used by teamsapp login)
  authority: 'https://login.microsoftonline.com/common',
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

// In-memory session storage (use Redis in production)
const sessions = new Map();

// Store pending device code requests with their status
const pendingDeviceCodeRequests = new Map();

// MSAL Client
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: CONFIG.authority,
  }
};
const pca = new msal.PublicClientApplication(msalConfig);

// ═══════════════════════════════════════════════════════════════
// AUTHENTICATION ENDPOINTS
// ═══════════════════════════════════════════════════════════════

/**
 * POST /api/auth/start
 * Start device code authentication flow using MSAL
 * This properly caches tokens for acquireTokenSilent to work
 */
app.post('/api/auth/start', async (req, res) => {
  try {
    const requestId = generateSessionId();

    // Create a promise that resolves when device code callback is invoked
    let resolveDeviceCode;
    const deviceCodePromise = new Promise((resolve) => {
      resolveDeviceCode = resolve;
    });

    // Use MSAL's built-in device code flow
    const deviceCodeRequest = {
      deviceCodeCallback: (response) => {
        // This callback receives the device code info
        const deviceCodeInfo = {
          userCode: response.userCode,
          verificationUri: response.verificationUri,
          message: response.message,
          expiresIn: response.expiresIn,
        };
        console.log('Device code generated:', response.userCode);
        resolveDeviceCode(deviceCodeInfo);
      },
      scopes: CONFIG.graphScopes,
    };

    // Start the device code flow (this will block until user authenticates)
    // We'll run it in the background and poll for completion
    const authPromise = pca.acquireTokenByDeviceCode(deviceCodeRequest);

    // Store the promise and track its completion
    const requestInfo = {
      promise: authPromise,
      completed: false,
      result: null,
      error: null,
      startTime: Date.now(),
    };

    // Track when the promise resolves
    authPromise
      .then((result) => {
        requestInfo.completed = true;
        requestInfo.result = result;
        console.log('Device code authentication completed for request:', requestId);
      })
      .catch((error) => {
        requestInfo.completed = true;
        requestInfo.error = error;
        console.error('Device code authentication failed for request:', requestId, error);
      });

    pendingDeviceCodeRequests.set(requestId, requestInfo);

    // Wait for the callback to be called with device code info
    const deviceCodeInfo = await Promise.race([
      deviceCodePromise,
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error('Timeout waiting for device code')), 5000)
      )
    ]);

    // Return device code info to client
    res.json({
      requestId: requestId,
      userCode: deviceCodeInfo.userCode,
      verificationUri: deviceCodeInfo.verificationUri,
      message: deviceCodeInfo.message,
      expiresIn: deviceCodeInfo.expiresIn,
    });

  } catch (error) {
    console.error('Auth start error:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * POST /api/auth/poll
 * Poll for authentication completion
 * Checks if MSAL's device code flow has completed
 */
app.post('/api/auth/poll', async (req, res) => {
  const { requestId } = req.body;

  try {
    const requestInfo = pendingDeviceCodeRequests.get(requestId);

    if (!requestInfo) {
      return res.status(400).json({ error: 'Invalid request ID' });
    }

    // Check if authentication has completed
    if (!requestInfo.completed) {
      // Still waiting for user to authenticate
      return res.json({ pending: true });
    }

    // Check if there was an error
    if (requestInfo.error) {
      pendingDeviceCodeRequests.delete(requestId);
      return res.status(500).json({ error: requestInfo.error.message });
    }

    // Authentication completed successfully!
    const tokenResponse = requestInfo.result;

    // MSAL automatically cached the tokens, so acquireTokenSilent will work
    // Get the account from the token response
    const account = tokenResponse.account;

    // Create new session with account info
    const sessionId = generateSessionId();
    sessions.set(sessionId, {
      account: account,
      createdAt: Date.now(),
    });

    // Clean up the pending request
    pendingDeviceCodeRequests.delete(requestId);

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
    console.error('Auth poll error:', error);
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

    const response = await pca.acquireTokenSilent(silentRequest);

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
// PROVISIONING ENDPOINTS
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

    const response = await pca.acquireTokenSilent(silentRequest);
    return response.accessToken;

  } catch (error) {
    console.error('Failed to acquire token silently:', error.message);
    throw error;
  }
}

function generateSessionId() {
  return Math.random().toString(36).substring(2) + Date.now().toString(36);
}

function decodeJwt(token) {
  if (!token) {
    throw new Error('Token is undefined or null');
  }

  const parts = token.split('.');
  if (parts.length !== 3) {
    throw new Error('Invalid JWT token format');
  }

  const base64Url = parts[1];
  const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
  const jsonPayload = decodeURIComponent(
    Buffer.from(base64, 'base64')
      .toString()
      .split('')
      .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
      .join('')
  );
  return JSON.parse(jsonPayload);
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
  console.log(`\n🚀 Bot Provisioner Backend running on http://localhost:${PORT}`);
  console.log(`\n📋 Available endpoints:`);
  console.log(`   POST /api/auth/start          - Start device code flow`);
  console.log(`   POST /api/auth/poll           - Poll for auth completion`);
  console.log(`   POST /api/provision/aad-app   - Create AAD app`);
  console.log(`   POST /api/provision/client-secret - Generate secret`);
  console.log(`   POST /api/provision/teams-app - Create Teams app`);
  console.log(`   POST /api/provision/bot       - Register bot`);
  console.log(`   POST /api/provision/complete  - Get final credentials\n`);
});
