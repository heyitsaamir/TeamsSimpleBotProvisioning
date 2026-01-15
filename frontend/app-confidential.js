// Configuration
const API_BASE = 'http://localhost:3003/api';

// State
let sessionId = null;
let credentials = {};

// ═══════════════════════════════════════════════════════════════
// PAGE LOAD
// ═══════════════════════════════════════════════════════════════

window.addEventListener('DOMContentLoaded', async () => {
  // Check if returning from admin consent
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('consent') === 'granted') {
    alert('Admin consent granted! Click "Check Scopes" to verify.');
    window.history.replaceState({}, document.title, window.location.pathname);
  }

  // Check if we have a session from redirect
  const storedSessionId = localStorage.getItem('sessionId');
  const storedUserInfo = localStorage.getItem('userInfo');

  if (storedSessionId && storedUserInfo) {
    sessionId = storedSessionId;
    const userInfo = JSON.parse(storedUserInfo);

    // Show authenticated state
    document.getElementById('user-email').textContent = userInfo.username;
    document.getElementById('tenant-id').textContent = userInfo.tenantId;
    document.getElementById('auth-status').style.display = 'block';

    // Auto-check scopes after authentication
    try {
      await checkAdminConsent();
    } catch (error) {
      // Session might be invalid (backend restarted), clear and show message
      if (error.message.includes('401') || error.message.includes('403')) {
        console.log('Session expired, clearing localStorage');
        localStorage.removeItem('sessionId');
        localStorage.removeItem('userInfo');
        sessionId = null;
        document.getElementById('auth-status').style.display = 'none';
        document.getElementById('scope-check-result').style.display = 'block';
        document.getElementById('scope-check-result').innerHTML = `
          <p style="color: orange;"><strong>Session expired.</strong> Please click "Check Scopes" to sign in again.</p>
        `;
      }
    }
  }
});

// ═══════════════════════════════════════════════════════════════
// CHECK SCOPES BUTTON
// ═══════════════════════════════════════════════════════════════

document.getElementById('btn-check-scopes').addEventListener('click', async () => {
  // If not authenticated, start OAuth flow first
  if (!sessionId) {
    try {
      const response = await fetch(`${API_BASE}/auth/start`, {
        method: 'GET',
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'Failed to start authentication');
      }

      // Redirect user to authorization URL
      window.location.href = data.authUrl;

    } catch (error) {
      alert('Authentication error: ' + error.message);
      console.error(error);
    }
    return;
  }

  // Already authenticated, check scopes
  await checkAdminConsent();
});

async function checkAdminConsent() {
  const resultDiv = document.getElementById('scope-check-result');
  const startProvisionBtn = document.getElementById('btn-start-provision');

  resultDiv.style.display = 'block';
  resultDiv.innerHTML = '<p>Checking scopes...</p>';

  try {
    const response = await fetch(`${API_BASE}/auth/check-consent`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId }),
    });

    const data = await response.json();

    if (!response.ok) {
      // If 401/403, session is invalid - clear localStorage
      if (response.status === 401 || response.status === 403) {
        localStorage.removeItem('sessionId');
        localStorage.removeItem('userInfo');
        sessionId = null;
        document.getElementById('auth-status').style.display = 'none';
        resultDiv.innerHTML = `<p style="color: orange;"><strong>Session expired.</strong> Please click "Check Scopes" to sign in again.</p>`;
        throw new Error(`${response.status}: ${data.error}`);
      }
      resultDiv.innerHTML = `<p style="color: red;"><strong>Error:</strong> ${data.error}</p>`;
      return;
    }

    if (data.hasConsent) {
      // All good! Admin has granted all permissions
      console.log('✓ Admin consent granted:', data.grantedScopes);

      resultDiv.innerHTML = `
        <p style="color: green;"><strong>✓ All required scopes are granted!</strong></p>
        <p>Granted scopes: ${data.grantedScopes.join(', ')}</p>
      `;

      // Enable Step 2
      document.getElementById('step-2').disabled = false;

    } else {
      // Missing admin consent - show warning
      console.warn('⚠ Missing permissions:', data.missingScopes);

      resultDiv.innerHTML = `
        <p style="color: orange;"><strong>⚠️ Admin Consent Required</strong></p>
        <p>Missing scopes: ${data.missingScopes.join(', ')}</p>
        <p><strong>Admin Consent URL:</strong></p>
        <p><input type="text" id="admin-consent-url-inline" value="${data.adminConsentUrl}" readonly style="width: 500px;"></p>
        <button id="btn-copy-consent-url-inline">Copy URL</button>
        <p><small>Send this URL to your admin, or click it if you're an admin: <a href="${data.adminConsentUrl}" target="_blank">Consent Now</a></small></p>
      `;

      // Add copy button handler
      document.getElementById('btn-copy-consent-url-inline').addEventListener('click', () => {
        const input = document.getElementById('admin-consent-url-inline');
        input.select();
        document.execCommand('copy');
        alert('Admin consent URL copied to clipboard!');
      });

      // Keep Step 2 disabled
      document.getElementById('step-2').disabled = true;
    }

  } catch (error) {
    console.error('Check consent error:', error);
    resultDiv.innerHTML = `<p style="color: red;"><strong>Error:</strong> ${error.message}</p>`;
  }
}

// ═══════════════════════════════════════════════════════════════
// STEP 2: CONFIGURATION
// ═══════════════════════════════════════════════════════════════

document.getElementById('btn-start-provision').addEventListener('click', async () => {
  const botName = document.getElementById('bot-name').value;
  const botEndpoint = document.getElementById('bot-endpoint').value;

  if (!botName || !botEndpoint) {
    alert('Please fill in all fields');
    return;
  }

  if (!botEndpoint.startsWith('https://')) {
    alert('Bot endpoint must be HTTPS');
    return;
  }

  // Disable step 2, enable step 3
  document.getElementById('step-2').disabled = true;
  document.getElementById('step-3').disabled = false;

  await runProvisioning(botName, botEndpoint);
});

// ═══════════════════════════════════════════════════════════════
// STEP 3: PROVISIONING (REUSED FROM app.js)
// ═══════════════════════════════════════════════════════════════

async function runProvisioning(botName, botEndpoint) {
  try {
    // Step 3.1: Create AAD App
    updateStatus('provision-aad', 'Creating...');
    const aadAppResponse = await fetch(`${API_BASE}/provision/aad-app`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, appName: botName }),
    });
    const aadApp = await aadAppResponse.json();

    if (!aadAppResponse.ok) {
      throw new Error(aadApp.error || 'Failed to create AAD app');
    }

    credentials.CLIENT_ID = aadApp.clientId;
    updateStatus('provision-aad', `✓ Created (${aadApp.clientId})`);

    // Step 3.2: Generate Client Secret
    updateStatus('provision-secret', 'Generating...');
    const secretResponse = await fetch(`${API_BASE}/provision/client-secret`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, objectId: aadApp.objectId }),
    });
    const secret = await secretResponse.json();

    if (!secretResponse.ok) {
      throw new Error(secret.error || 'Failed to generate client secret');
    }

    credentials.CLIENT_SECRET = secret.clientSecret;
    updateStatus('provision-secret', '✓ Generated');

    // Step 3.3: Create Teams App
    updateStatus('provision-teams', 'Creating...');
    const manifest = {
      $schema: 'https://developer.microsoft.com/json-schemas/teams/v1.16/MicrosoftTeams.schema.json',
      manifestVersion: '1.16',
      version: '1.0.0',
      id: aadApp.clientId,
      packageName: 'com.example.bot',
      developer: {
        name: 'Developer',
        websiteUrl: 'https://www.example.com',
        privacyUrl: 'https://www.example.com/privacy',
        termsOfUseUrl: 'https://www.example.com/terms'
      },
      name: {
        short: botName,
        full: botName
      },
      description: {
        short: botName,
        full: botName
      },
      icons: {
        color: 'color.png',
        outline: 'outline.png'
      },
      accentColor: '#FFFFFF',
      bots: [
        {
          botId: aadApp.clientId,
          scopes: ['personal', 'team', 'groupchat'],
          supportsFiles: false,
          isNotificationOnly: false
        }
      ],
      permissions: ['identity', 'messageTeamMembers'],
      validDomains: []
    };

    const teamsAppResponse = await fetch(`${API_BASE}/provision/teams-app`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, manifest }),
    });
    const teamsApp = await teamsAppResponse.json();

    if (!teamsAppResponse.ok) {
      throw new Error(teamsApp.error || 'Failed to create Teams app');
    }

    credentials.TEAMS_APP_ID = teamsApp.teamsAppId;
    updateStatus('provision-teams', `✓ Created (${teamsApp.teamsAppId})`);

    // Step 3.4: Register Bot
    updateStatus('provision-bot', 'Registering...');
    const botResponse = await fetch(`${API_BASE}/provision/bot`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId,
        botId: aadApp.clientId,
        botName: botName,
        messagingEndpoint: botEndpoint
      }),
    });
    const botResult = await botResponse.json();

    if (!botResponse.ok) {
      throw new Error(botResult.error || 'Failed to register bot');
    }

    credentials.BOT_ENDPOINT = botEndpoint;
    updateStatus('provision-bot', '✓ Registered');

    // Complete
    const complete = await fetch(`${API_BASE}/provision/complete`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, credentials }),
    }).then(r => r.json());

    credentials.TENANT_ID = complete.credentials.TENANT_ID;

    // Show credentials
    document.getElementById('step-3').disabled = true;
    document.getElementById('step-4').disabled = false;
    displayCredentials();

  } catch (error) {
    console.error('Provisioning error:', error);
    document.getElementById('provision-error-message').textContent = error.message;
    document.getElementById('provision-error').style.display = 'block';
  }
}

function updateStatus(elementId, status) {
  const el = document.getElementById(elementId);
  if (el) {
    el.querySelector('span').textContent = status;
  }
}

// ═══════════════════════════════════════════════════════════════
// STEP 4: COMPLETE
// ═══════════════════════════════════════════════════════════════

function displayCredentials() {
  document.getElementById('cred-client-id').textContent = credentials.CLIENT_ID;
  document.getElementById('cred-client-secret').textContent = credentials.CLIENT_SECRET;
  document.getElementById('cred-tenant-id').textContent = credentials.TENANT_ID;
  document.getElementById('cred-teams-app-id').textContent = credentials.TEAMS_APP_ID;
  document.getElementById('cred-bot-endpoint').textContent = credentials.BOT_ENDPOINT;

  const envContent = `CLIENT_ID=${credentials.CLIENT_ID}
CLIENT_SECRET=${credentials.CLIENT_SECRET}
TENANT_ID=${credentials.TENANT_ID}
TEAMS_APP_ID=${credentials.TEAMS_APP_ID}
BOT_ENDPOINT=${credentials.BOT_ENDPOINT}`;

  document.getElementById('env-file-content').textContent = envContent;
}

// Copy buttons
document.addEventListener('click', (e) => {
  if (e.target.classList.contains('btn-copy')) {
    const targetId = e.target.getAttribute('data-copy');
    const text = document.getElementById(targetId).textContent;
    navigator.clipboard.writeText(text);
    e.target.textContent = 'Copied!';
    setTimeout(() => {
      e.target.textContent = 'Copy';
    }, 2000);
  }
});

// Download .env
document.getElementById('btn-download-env').addEventListener('click', () => {
  const envContent = document.getElementById('env-file-content').textContent;
  const blob = new Blob([envContent], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '.env';
  a.click();
});

// Start over
document.getElementById('btn-start-over').addEventListener('click', () => {
  localStorage.removeItem('sessionId');
  localStorage.removeItem('userInfo');
  window.location.reload();
});

// Retry provisioning
document.getElementById('btn-retry').addEventListener('click', () => {
  document.getElementById('provision-error').style.display = 'none';
  const botName = document.getElementById('bot-name').value;
  const botEndpoint = document.getElementById('bot-endpoint').value;
  runProvisioning(botName, botEndpoint);
});
