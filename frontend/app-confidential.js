// ===== CONFIGURATION =====
// Set this to true if backend is also tunneled, false if backend is on localhost
const BACKEND_TUNNELED = false; // Change to true when backend tunnel is ready

// Configure backend URL
const API_BASE = BACKEND_TUNNELED
  ? 'https://3hvfdfhp-3002.usw2.devtunnels.ms' // Backend tunnel URL
  : 'http://localhost:3002'; // Backend on localhost (may cause mixed content error on HTTPS)

// State
let sessionId = null;
let credentials = {};

// Check if returning from OAuth redirect
window.addEventListener('DOMContentLoaded', () => {
  // Check if we have a session from redirect
  const storedSessionId = localStorage.getItem('sessionId');
  const storedUserInfo = localStorage.getItem('userInfo');

  if (storedSessionId && storedUserInfo) {
    sessionId = storedSessionId;
    const userInfo = JSON.parse(storedUserInfo);

    // Show authenticated state
    document.getElementById('user-email').textContent = userInfo.username;
    document.getElementById('tenant-id').textContent = userInfo.tenantId;
    document.getElementById('session-id').textContent = sessionId;
    document.getElementById('auth-success').style.display = 'block';
    document.getElementById('btn-auth').style.display = 'none';

    // Enable step 2
    document.getElementById('step-2').disabled = false;
  }
});

// Step 1: Authentication
document.getElementById('btn-auth').addEventListener('click', async () => {
  try {
    const response = await fetch(`${API_BASE}/api/auth/start`, {
      method: 'GET',
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || 'Failed to start authentication');
    }

    // Store state for verification (optional)
    localStorage.setItem('authState', data.state);

    // Redirect user to authorization URL
    window.location.href = data.authUrl;

  } catch (error) {
    alert('Authentication error: ' + error.message);
    console.error(error);
  }
});

// Step 2: Start provisioning
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

async function runProvisioning(botName, botEndpoint) {
  try {
    // Step 3.1: Create AAD App
    updateStatus('provision-aad', 'Creating...');
    const aadApp = await fetch(`${API_BASE}/api/provision/aad-app`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, appName: botName }),
    }).then(r => r.json());

    credentials.CLIENT_ID = aadApp.clientId;
    updateStatus('provision-aad', `✓ Created (${aadApp.clientId})`);

    // Step 3.2: Generate Client Secret
    updateStatus('provision-secret', 'Generating...');
    const secret = await fetch(`${API_BASE}/api/provision/client-secret`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, objectId: aadApp.objectId }),
    }).then(r => r.json());

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

    const teamsApp = await fetch(`${API_BASE}/api/provision/teams-app`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, manifest }),
    }).then(r => r.json());

    credentials.TEAMS_APP_ID = teamsApp.teamsAppId;
    updateStatus('provision-teams', `✓ Created (${teamsApp.teamsAppId})`);

    // Step 3.4: Register Bot
    updateStatus('provision-bot', 'Registering...');
    await fetch(`${API_BASE}/api/provision/bot`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId,
        botId: aadApp.clientId,
        botName: botName,
        messagingEndpoint: botEndpoint
      }),
    }).then(r => r.json());

    credentials.BOT_ENDPOINT = botEndpoint;
    updateStatus('provision-bot', '✓ Registered');

    // Complete
    const complete = await fetch(`${API_BASE}/api/provision/complete`, {
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
  localStorage.removeItem('authState');
  window.location.reload();
});

// Retry provisioning
document.getElementById('btn-retry').addEventListener('click', () => {
  document.getElementById('provision-error').style.display = 'none';
  const botName = document.getElementById('bot-name').value;
  const botEndpoint = document.getElementById('bot-endpoint').value;
  runProvisioning(botName, botEndpoint);
});
