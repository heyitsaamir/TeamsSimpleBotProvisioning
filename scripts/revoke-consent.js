/**
 * Helper script to revoke admin consent for testing
 * Usage: node scripts/revoke-consent.js
 */

const msal = require('@azure/msal-node');
const axios = require('axios');

const CONFIG = {
    clientId: '2a098349-9ecc-463f-a053-d5675e10deeb',
    authority: 'https://login.microsoftonline.com/common',
};

async function revokeConsent() {
    console.log('🔑 Starting device code authentication...\n');

    const pca = new msal.PublicClientApplication({
        auth: {
            clientId: CONFIG.clientId,
            authority: CONFIG.authority,
        }
    });

    // Get token with admin scopes
    const deviceCodeRequest = {
        deviceCodeCallback: (response) => {
            console.log('📱 Please authenticate:\n');
            console.log(`   URL: ${response.verificationUri}`);
            console.log(`   Code: ${response.userCode}\n`);
        },
        scopes: [
            'https://graph.microsoft.com/Application.ReadWrite.All',
            'https://graph.microsoft.com/Directory.ReadWrite.All'
        ],
    };

    const authResult = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    const token = authResult.accessToken;

    console.log('✅ Authenticated\n');

    // Find the service principal
    console.log('🔍 Finding service principal...');
    const spResponse = await axios.get(
        'https://graph.microsoft.com/v1.0/servicePrincipals',
        {
            params: {
                '$filter': `appId eq '${CONFIG.clientId}'`
            },
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            }
        }
    );

    const servicePrincipal = spResponse.data.value[0];

    if (!servicePrincipal) {
        console.log('❌ Service principal not found. No consent to revoke.');
        return;
    }

    console.log(`✓ Found service principal: ${servicePrincipal.displayName}`);
    console.log(`   ID: ${servicePrincipal.id}\n`);

    // Get all OAuth2 permission grants
    console.log('🔍 Finding permission grants...');
    const grantsResponse = await axios.get(
        'https://graph.microsoft.com/v1.0/oauth2PermissionGrants',
        {
            params: {
                '$filter': `clientId eq '${servicePrincipal.id}'`
            },
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            }
        }
    );

    const grants = grantsResponse.data.value;

    if (grants.length === 0) {
        console.log('❌ No permission grants found.\n');
    } else {
        console.log(`✓ Found ${grants.length} permission grant(s)\n`);

        // Delete each grant
        for (const grant of grants) {
            console.log(`Revoking: ${grant.scope || '(no scopes)'}`);
            try {
                await axios.delete(
                    `https://graph.microsoft.com/v1.0/oauth2PermissionGrants/${grant.id}`,
                    {
                        headers: {
                            'Authorization': `Bearer ${token}`,
                        }
                    }
                );
                console.log('  ✓ Revoked\n');
            } catch (error) {
                console.log(`  ❌ Failed: ${error.response?.data?.error?.message || error.message}\n`);
            }
        }
    }

    // Optional: Delete the service principal entirely
    console.log('❓ Do you want to DELETE the service principal entirely?');
    console.log('   This removes all consent for all users in your tenant.');
    console.log('   (Press Ctrl+C to skip, or wait 5 seconds to delete...)\n');

    await new Promise(resolve => setTimeout(resolve, 5000));

    console.log('🗑️  Deleting service principal...');
    try {
        await axios.delete(
            `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipal.id}`,
            {
                headers: {
                    'Authorization': `Bearer ${token}`,
                }
            }
        );
        console.log('✅ Service principal deleted\n');
    } catch (error) {
        console.log(`❌ Failed to delete: ${error.response?.data?.error?.message || error.message}\n`);
    }

    console.log('✅ Done! Consent has been revoked.');
    console.log('   Next sign-in will require fresh consent.\n');
}

revokeConsent().catch(error => {
    console.error('Error:', error.message);
    process.exit(1);
});
