/**
 * Bot Provisioner - Reference Frontend Implementation
 *
 * This file demonstrates how to build a frontend that:
 * 1. Initiates OAuth authentication flow
 * 2. Checks if admin consent has been granted
 * 3. Orchestrates bot provisioning via backend API calls
 * 4. Displays generated credentials and Teams installation link
 */

// Configuration - adjust to match your backend
const API_BASE = 'http://localhost:3003/api';

/**
 * Session State - How We Track the User Across Requests
 *
 * The sessionId acts as our way to maintain state for the authenticated user:
 * - After OAuth callback, backend creates a session and returns sessionId
 * - We store sessionId in localStorage (persists across page reloads)
 * - We send sessionId with EVERY backend API request
 * - Backend uses sessionId to look up the user's account object
 * - This allows backend to call acquireTokenSilent() for that user
 *
 * What's stored here:
 * - sessionId: Random identifier for the session
 * - userInfo: Basic user info (username, tenantId) for display only
 *
 * What's NOT stored:
 * - Tokens (backend manages these via MSAL)
 * - Secrets (never in frontend)
 */
let sessionId = localStorage.getItem('sessionId');
let userInfo = null;

// ═══════════════════════════════════════════════════════════════
// INITIALIZATION
// ═══════════════════════════════════════════════════════════════

/**
 * Initialize the application on page load.
 *
 * If user has an existing session (from localStorage), restore it.
 * Otherwise, wait for user to click "Check Scopes" button.
 */
window.addEventListener('DOMContentLoaded', () => {
    // Restore session if exists
    if (sessionId) {
        const storedUserInfo = localStorage.getItem('userInfo');
        if (storedUserInfo) {
            userInfo = JSON.parse(storedUserInfo);
            showAuthStatus();
            enableSideloadingCheck();
        }
    }

    // Attach event listeners
    document.getElementById('btn-check-scopes').addEventListener('click', checkScopes);
    document.getElementById('btn-check-sideloading').addEventListener('click', checkSideloading);
    document.getElementById('btn-provision').addEventListener('click', startProvisioning);
});

// ═══════════════════════════════════════════════════════════════
// AUTHENTICATION & SCOPE CHECKING
// ═══════════════════════════════════════════════════════════════

/**
 * Initiates scope check, which triggers authentication if needed.
 *
 * Flow:
 * 1. If not authenticated, start OAuth flow
 * 2. After authentication, check if required scopes are granted
 * 3. Display granted/missing scopes
 * 4. Show admin consent URL if scopes are missing
 */
async function checkScopes() {
    const resultDiv = document.getElementById('scope-check-result');
    resultDiv.style.display = 'block';
    resultDiv.innerHTML = '<p>Checking permissions...</p>';

    try {
        // If no session, need to authenticate first
        if (!sessionId) {
            await authenticate();
        }

        // Check consent status
        const response = await fetch(`${API_BASE}/auth/check-consent`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ sessionId }),
        });

        if (!response.ok) {
            // Session might be invalid - clear and retry
            if (response.status === 401 || response.status === 403) {
                clearSession();
                resultDiv.innerHTML = '<p class="warning">Session expired. Please try again.</p>';
                return;
            }
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();

        if (data.hasConsent) {
            // All scopes granted!
            resultDiv.innerHTML = `
                <div class="success">
                    <p><strong>✓ All required permissions granted</strong></p>
                    <p>Granted scopes: ${data.grantedScopes.join(', ')}</p>
                </div>
            `;

            // Enable provisioning
            document.getElementById('step-2').classList.remove('hidden');
            enableSideloadingCheck();

        } else {
            // Some scopes missing - show admin consent URL
            resultDiv.innerHTML = `
                <div class="warning">
                    <p><strong>⚠️ Admin consent required</strong></p>
                    <p><strong>Granted scopes:</strong> ${data.grantedScopes.length > 0 ? data.grantedScopes.join(', ') : 'None'}</p>
                    <p><strong>Missing scopes:</strong> ${data.missingScopes.join(', ')}</p>
                    <p>Your tenant admin must grant consent for these permissions.</p>
                    <p><strong>Admin consent URL:</strong></p>
                    <p><a href="${data.adminConsentUrl}" target="_blank">${data.adminConsentUrl}</a></p>
                    <p style="margin-top: 10px; font-size: 0.9em;">After admin grants consent, click "Check Scopes" again.</p>
                </div>
            `;
        }

    } catch (error) {
        console.error('Scope check error:', error);
        resultDiv.innerHTML = `<div class="error"><strong>Error:</strong> ${error.message}</div>`;
    }
}

/**
 * Authenticates the user via OAuth authorization code flow.
 *
 * Flow:
 * 1. Get authorization URL from backend
 * 2. Redirect user to Azure AD login
 * 3. Azure AD redirects back to redirect.html with authorization code
 * 4. redirect.html exchanges code for tokens and redirects back here
 */
async function authenticate() {
    try {
        // Get authorization URL from backend
        const response = await fetch(`${API_BASE}/auth/start`);
        const data = await response.json();

        // Redirect to Azure AD login
        window.location.href = data.authUrl;

    } catch (error) {
        console.error('Authentication error:', error);
        alert('Failed to start authentication: ' + error.message);
    }
}

/**
 * Displays authenticated user information.
 */
function showAuthStatus() {
    const statusDiv = document.getElementById('auth-status');
    const usernameSpan = document.getElementById('username');

    usernameSpan.textContent = userInfo.username;
    statusDiv.classList.remove('hidden');
}

/**
 * Enables the sideloading check button after authentication.
 */
function enableSideloadingCheck() {
    document.getElementById('btn-check-sideloading').disabled = false;
}

/**
 * Clears session data from localStorage.
 */
function clearSession() {
    localStorage.removeItem('sessionId');
    localStorage.removeItem('userInfo');
    sessionId = null;
    userInfo = null;
}

// ═══════════════════════════════════════════════════════════════
// SIDELOADING CHECK
// ═══════════════════════════════════════════════════════════════

/**
 * Checks if the user's tenant allows custom app sideloading.
 *
 * This is required for users to install the provisioned bot.
 * If sideloading is disabled, the tenant admin must enable it.
 */
async function checkSideloading() {
    const resultDiv = document.getElementById('sideloading-check-result');
    resultDiv.style.display = 'block';
    resultDiv.innerHTML = '<p>Checking sideloading status...</p>';

    try {
        const response = await fetch(`${API_BASE}/check-sideloading`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ sessionId }),
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();

        if (data.status === 'enabled') {
            resultDiv.innerHTML = `
                <div class="success">
                    <p><strong>✓ Sideloading is ENABLED</strong></p>
                    <p>Your organization allows custom app sideloading in Microsoft Teams.</p>
                </div>
            `;
        } else if (data.status === 'disabled') {
            resultDiv.innerHTML = `
                <div class="error">
                    <p><strong>✗ Sideloading is DISABLED</strong></p>
                    <p>Your organization does not allow custom app sideloading.</p>
                    <p>Contact your Teams administrator to enable it in the Teams admin center.</p>
                </div>
            `;
        } else {
            resultDiv.innerHTML = `
                <div class="warning">
                    <p><strong>⚠️ Sideloading status unknown</strong></p>
                    <p>Could not determine sideloading status.</p>
                </div>
            `;
        }

    } catch (error) {
        console.error('Sideloading check error:', error);
        resultDiv.innerHTML = `<div class="error"><strong>Error:</strong> ${error.message}</div>`;
    }
}

// ═══════════════════════════════════════════════════════════════
// BOT PROVISIONING
// ═══════════════════════════════════════════════════════════════

/**
 * Orchestrates the complete bot provisioning flow.
 *
 * Steps:
 * 1. Create Azure AD app registration
 * 2. Generate client secret
 * 3. Create Teams app package
 * 4. Register bot with Bot Framework
 * 5. Display credentials and installation link
 */
async function startProvisioning() {
    const progressDiv = document.getElementById('provision-progress');
    const provisionBtn = document.getElementById('btn-provision');

    progressDiv.style.display = 'block';
    progressDiv.innerHTML = '';
    provisionBtn.disabled = true;

    // Get bot details from form
    const botName = document.getElementById('bot-name').value;
    const botEndpoint = document.getElementById('bot-endpoint').value;

    if (!botName || !botEndpoint) {
        alert('Please fill in all bot details');
        provisionBtn.disabled = false;
        return;
    }

    try {
        // Step 1: Create Azure AD app
        progressDiv.innerHTML += '<p>📝 Creating Azure AD app registration...</p>';
        const aadApp = await createAadApp(botName);
        const clientId = aadApp.clientId;
        const appRegistrationId = aadApp.appRegistrationId;
        progressDiv.innerHTML += `<p class="success">✓ Created AAD app: ${clientId}</p>`;

        // Step 2: Generate client secret
        progressDiv.innerHTML += '<p>🔑 Generating client secret...</p>';
        const secret = await createClientSecret(appRegistrationId);
        const clientSecret = secret.clientSecret;
        progressDiv.innerHTML += '<p class="success">✓ Generated client secret</p>';

        // Step 3: Create Teams app
        progressDiv.innerHTML += '<p>📦 Creating Teams app package...</p>';
        const teamsApp = await createTeamsApp(clientId, botName);
        const teamsAppId = teamsApp.teamsAppId;
        progressDiv.innerHTML += `<p class="success">✓ Created Teams app: ${teamsAppId}</p>`;

        // Step 4: Register bot
        progressDiv.innerHTML += '<p>🤖 Registering bot with Bot Framework...</p>';
        await registerBot(clientId, botName, botEndpoint);
        progressDiv.innerHTML += '<p class="success">✓ Registered bot</p>';

        // Show results
        displayResults({
            botId: clientId,
            botPassword: clientSecret,
            teamsAppId: teamsAppId,
            tenantId: userInfo.tenantId,
        });

    } catch (error) {
        console.error('Provisioning error:', error);
        progressDiv.innerHTML += `<div class="error"><strong>Error:</strong> ${error.message}</div>`;
        provisionBtn.disabled = false;
    }
}

/**
 * Step 1: Creates an Azure AD app registration.
 */
async function createAadApp(appName) {
    const response = await fetch(`${API_BASE}/provision/aad-app`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            sessionId: sessionId,
            appName: appName,
        }),
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(`AAD app creation failed: ${error.error}`);
    }

    return await response.json();
}

/**
 * Step 2: Generates a client secret for the AAD app.
 */
async function createClientSecret(appRegistrationId) {
    const response = await fetch(`${API_BASE}/provision/client-secret`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            sessionId: sessionId,
            appRegistrationId: appRegistrationId,
        }),
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(`Secret generation failed: ${error.error}`);
    }

    return await response.json();
}

/**
 * Step 3: Creates a Teams app package and uploads to TDP.
 */
async function createTeamsApp(clientId, botName) {
    // Create Teams app manifest
    const manifest = {
        "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.16/MicrosoftTeams.schema.json",
        "manifestVersion": "1.16",
        "version": "1.0.0",
        "id": clientId,
        "packageName": `com.teams.${clientId}`,
        "developer": {
            "name": "Bot Developer",
            "websiteUrl": "https://www.example.com",
            "privacyUrl": "https://www.example.com/privacy",
            "termsOfUseUrl": "https://www.example.com/terms"
        },
        "icons": {
            "color": "color.png",
            "outline": "outline.png"
        },
        "name": {
            "short": botName,
            "full": botName
        },
        "description": {
            "short": botName,
            "full": botName
        },
        "accentColor": "#FFFFFF",
        "bots": [
            {
                "botId": clientId,
                "scopes": ["personal", "team", "groupchat"],
                "supportsFiles": false,
                "isNotificationOnly": false
            }
        ],
        "permissions": ["identity", "messageTeamMembers"],
        "validDomains": []
    };

    const response = await fetch(`${API_BASE}/provision/teams-app`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            sessionId: sessionId,
            manifest: manifest,
        }),
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(`Teams app creation failed: ${error.error}`);
    }

    return await response.json();
}

/**
 * Step 4: Registers the bot with Bot Framework.
 */
async function registerBot(botId, botName, messagingEndpoint) {
    const response = await fetch(`${API_BASE}/provision/bot`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            sessionId: sessionId,
            botId: botId,
            botName: botName,
            messagingEndpoint: messagingEndpoint,
        }),
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(`Bot registration failed: ${error.error}`);
    }

    return await response.json();
}

// ═══════════════════════════════════════════════════════════════
// RESULTS DISPLAY
// ═══════════════════════════════════════════════════════════════

/**
 * Displays the generated credentials and Teams installation link.
 */
function displayResults(credentials) {
    // Show credentials
    const credentialsText = `BOT_ID=${credentials.botId}
BOT_PASSWORD=${credentials.botPassword}
TEAMS_APP_ID=${credentials.teamsAppId}
TENANT_ID=${credentials.tenantId}`;

    document.getElementById('credentials').textContent = credentialsText;

    // Generate Teams deep link
    const deepLink = `https://teams.microsoft.com/l/app/${credentials.teamsAppId}?installAppPackage=true&webjoin=true&appTenantId=${credentials.tenantId}&login_hint=${encodeURIComponent(userInfo.username)}`;

    const deepLinkElement = document.getElementById('teams-deep-link');
    deepLinkElement.href = deepLink;
    deepLinkElement.textContent = 'Install Bot in Microsoft Teams';

    // Show results section
    document.getElementById('step-3').classList.remove('hidden');

    // Scroll to results
    document.getElementById('step-3').scrollIntoView({ behavior: 'smooth' });
}
