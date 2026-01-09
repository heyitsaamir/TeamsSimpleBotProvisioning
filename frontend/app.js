// Configuration
const API_BASE = 'http://localhost:3001/api';

// Application State
const state = {
  sessionId: null,
  authenticated: false,
  config: {
    botName: '',
    botEndpoint: '',
  },
  credentials: {
    clientId: '',
    clientSecret: '',
    tenantId: '',
    teamsAppId: '',
    botEndpoint: '',
  },
};

// ═══════════════════════════════════════════════════════════════
// UI HELPERS
// ═══════════════════════════════════════════════════════════════

function enableStep(stepNumber) {
  document.getElementById(`step-${stepNumber}`).disabled = false;
}

function disableStep(stepNumber) {
  document.getElementById(`step-${stepNumber}`).disabled = true;
}

function updateProvisionStep(stepId, message) {
  const element = document.getElementById(stepId);
  element.querySelector('span').textContent = message;
}

// ═══════════════════════════════════════════════════════════════
// STEP 1: AUTHENTICATION
// ═══════════════════════════════════════════════════════════════

document.getElementById('btn-auth').addEventListener('click', async () => {
  try {
    console.log('Starting authentication with MSAL...');

    const response = await fetch(`${API_BASE}/auth/start`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || 'Failed to start authentication');
    }

    const data = await response.json();
    console.log('Device code received:', data.userCode);
    console.log('Request ID:', data.requestId);
    console.log('MSAL will cache tokens for acquireTokenSilent');

    document.getElementById('auth-url').href = data.verificationUri;
    document.getElementById('user-code').textContent = data.userCode;
    document.getElementById('auth-device-code').style.display = 'block';
    document.getElementById('btn-auth').disabled = true;

    pollForAuth(data.requestId);

  } catch (error) {
    console.error('Authentication error:', error);
    alert('Error: ' + error.message);
  }
});

async function pollForAuth(requestId) {
  const poll = async () => {
    try {
      const response = await fetch(`${API_BASE}/auth/poll`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          requestId
        })
      });

      const data = await response.json();

      if (data.pending) {
        setTimeout(poll, 3000);
      } else if (data.success) {
        state.sessionId = data.sessionId;
        state.authenticated = true;

        document.getElementById('auth-device-code').style.display = 'none';
        document.getElementById('auth-success').style.display = 'block';
        document.getElementById('user-email').textContent = data.userInfo.username;
        document.getElementById('tenant-id').textContent = data.userInfo.tenantId;
        state.credentials.tenantId = data.userInfo.tenantId;

        console.log('Authentication complete! MSAL cached tokens for acquireTokenSilent.');

        // Enable Step 2 immediately - no second auth needed!
        enableStep(2);
      } else {
        throw new Error(data.error || 'Authentication failed');
      }
    } catch (error) {
      alert('Error: ' + error.message);
      document.getElementById('btn-auth').disabled = false;
    }
  };

  poll();
}

// ═══════════════════════════════════════════════════════════════
// STEP 2: CONFIGURATION
// ═══════════════════════════════════════════════════════════════

document.getElementById('btn-start-provision').addEventListener('click', () => {
  const botName = document.getElementById('bot-name').value;
  const botEndpoint = document.getElementById('bot-endpoint').value;

  if (!botName || !botEndpoint) {
    alert('Please fill in all required fields');
    return;
  }

  if (!botEndpoint.startsWith('https://')) {
    alert('Bot endpoint must be an HTTPS URL');
    return;
  }

  state.config.botName = botName;
  state.config.botEndpoint = botEndpoint;

  // Disable config step, enable provisioning step
  disableStep(2);
  enableStep(3);

  startProvisioning();
});

// ═══════════════════════════════════════════════════════════════
// STEP 3: PROVISIONING
// ═══════════════════════════════════════════════════════════════

async function startProvisioning() {
  try {
    // Step 1: Create AAD App
    updateProvisionStep('provision-aad', 'In progress...');
    const aadApp = await createAadApp();
    state.credentials.clientId = aadApp.clientId;
    updateProvisionStep('provision-aad', '✓ Done! ' + aadApp.clientId);

    // Step 2: Generate Secret
    updateProvisionStep('provision-secret', 'In progress...');
    const secret = await generateClientSecret(aadApp.objectId);
    state.credentials.clientSecret = secret.clientSecret;
    updateProvisionStep('provision-secret', '✓ Done!');

    // Step 3: Create Teams App
    updateProvisionStep('provision-teams', 'In progress...');
    const teamsApp = await createTeamsApp();
    state.credentials.teamsAppId = teamsApp.teamsAppId;
    updateProvisionStep('provision-teams', '✓ Done! ' + teamsApp.teamsAppId);

    // Step 4: Register Bot
    updateProvisionStep('provision-bot', 'In progress...');
    await registerBot(aadApp.clientId);
    state.credentials.botEndpoint = `${state.config.botEndpoint}/api/messages`;
    updateProvisionStep('provision-bot', '✓ Done!');

    // Success! Enable Step 4
    setTimeout(() => {
      enableStep(4);
      displayCredentials();
    }, 500);

  } catch (error) {
    console.error('Provisioning error:', error);
    document.getElementById('provision-error').style.display = 'block';
    document.getElementById('provision-error-message').textContent = error.message;
  }
}

async function createAadApp() {
  const response = await fetch(`${API_BASE}/provision/aad-app`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      sessionId: state.sessionId,
      appName: state.config.botName
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error || 'Failed to create AAD app');
  }

  return await response.json();
}

async function generateClientSecret(objectId) {
  const response = await fetch(`${API_BASE}/provision/client-secret`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      sessionId: state.sessionId,
      objectId: objectId
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error || 'Failed to generate client secret');
  }

  return await response.json();
}

async function createTeamsApp() {
  const manifest = {
    "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.16/MicrosoftTeams.schema.json",
    "manifestVersion": "1.16",
    "version": "1.0.0",
    "id": generateUuid(),
    "packageName": "com.example.bot",
    "developer": {
      "name": "Developer",
      "websiteUrl": "https://www.example.com",
      "privacyUrl": "https://www.example.com/privacy",
      "termsOfUseUrl": "https://www.example.com/terms"
    },
    "icons": {
      "color": "color.png",
      "outline": "outline.png"
    },
    "name": {
      "short": state.config.botName,
      "full": state.config.botName
    },
    "description": {
      "short": state.config.botName,
      "full": state.config.botName
    },
    "accentColor": "#FFFFFF",
    "bots": [{
      "botId": state.credentials.clientId,
      "scopes": ["personal", "team", "groupchat"],
      "supportsFiles": false,
      "isNotificationOnly": false
    }],
    "permissions": ["identity", "messageTeamMembers"],
    "validDomains": [new URL(state.config.botEndpoint).hostname]
  };

  const response = await fetch(`${API_BASE}/provision/teams-app`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      sessionId: state.sessionId,
      manifest: manifest
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error || 'Failed to create Teams app');
  }

  return await response.json();
}

async function registerBot(botId) {
  const response = await fetch(`${API_BASE}/provision/bot`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      sessionId: state.sessionId,
      botId: botId,
      botName: state.config.botName,
      messagingEndpoint: `${state.config.botEndpoint}/api/messages`
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error || 'Failed to register bot');
  }

  return await response.json();
}

document.getElementById('btn-retry').addEventListener('click', () => {
  document.getElementById('provision-error').style.display = 'none';
  updateProvisionStep('provision-aad', 'Waiting...');
  updateProvisionStep('provision-secret', 'Waiting...');
  updateProvisionStep('provision-teams', 'Waiting...');
  updateProvisionStep('provision-bot', 'Waiting...');
  startProvisioning();
});

// ═══════════════════════════════════════════════════════════════
// STEP 4: COMPLETE
// ═══════════════════════════════════════════════════════════════

function displayCredentials() {
  document.getElementById('cred-client-id').textContent = state.credentials.clientId;
  document.getElementById('cred-client-secret').textContent = state.credentials.clientSecret;
  document.getElementById('cred-tenant-id').textContent = state.credentials.tenantId;
  document.getElementById('cred-teams-app-id').textContent = state.credentials.teamsAppId;
  document.getElementById('cred-bot-endpoint').textContent = state.credentials.botEndpoint;

  // Generate Teams deep link and enable button
  const teamsDeepLink = `https://teams.microsoft.com/l/app/${state.credentials.teamsAppId}?installAppPackage=true&source=developerportal`;
  const deepLinkButton = document.getElementById('teams-deep-link-btn');

  if (deepLinkButton) {
    console.log('Enabling Teams deep link button for app:', state.credentials.teamsAppId);
    console.log('Button found:', deepLinkButton);
    console.log('Current disabled state:', deepLinkButton.disabled);

    deepLinkButton.textContent = '🚀 Open in Microsoft Teams';
    deepLinkButton.disabled = false;
    deepLinkButton.removeAttribute('disabled'); // Extra: explicitly remove attribute

    console.log('New disabled state:', deepLinkButton.disabled);

    // Add click handler to open in new window
    deepLinkButton.onclick = () => {
      console.log('Opening Teams deep link:', teamsDeepLink);
      window.open(teamsDeepLink, '_blank', 'noopener,noreferrer');
    };
  } else {
    console.error('Teams deep link button not found!');
  }

  const envContent = `BOT_ID=${state.credentials.clientId}
CLIENT_ID=${state.credentials.clientId}
CLIENT_SECRET=${state.credentials.clientSecret}
TENANT_ID=${state.credentials.tenantId}
TEAMS_APP_ID=${state.credentials.teamsAppId}
BOT_ENDPOINT=${state.config.botEndpoint}
MESSAGING_ENDPOINT=${state.credentials.botEndpoint}`;

  document.getElementById('env-file-content').textContent = envContent;
}

// Copy buttons
document.querySelectorAll('.btn-copy').forEach(btn => {
  btn.addEventListener('click', (e) => {
    const targetId = e.target.dataset.copy;
    const text = document.getElementById(targetId).textContent;
    navigator.clipboard.writeText(text);
    e.target.textContent = 'Copied!';
    setTimeout(() => {
      e.target.textContent = 'Copy';
    }, 2000);
  });
});

// Download .env file
document.getElementById('btn-download-env').addEventListener('click', () => {
  const content = document.getElementById('env-file-content').textContent;
  const blob = new Blob([content], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '.env';
  a.click();
  URL.revokeObjectURL(url);
});

// Start over
document.getElementById('btn-start-over').addEventListener('click', () => {
  if (confirm('Are you sure you want to start over?')) {
    location.reload();
  }
});

// Helper
function generateUuid() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}
