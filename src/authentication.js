// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

const getAccessToken = async function () {
    // Create a config.powerBI variable that store credentials from config.powerBI.json
    const config = require(__dirname + "/../config/custom-environment-variables.json");

    // Use MSAL.js for authentication
    const msal = require("@azure/msal-node");

    const msalConfig = {
        auth: {
            clientId: config.powerBI.clientId,
            authority: `${config.powerBI.authorityUrl}${config.powerBI.tenantId}`,
        }
    };

    // Check for the MasterUser Authentication
    if (config.powerBI.authenticationMode.toLowerCase() === "masteruser") {
        const clientApplication = new msal.PublicClientApplication(msalConfig);

        const usernamePasswordRequest = {
            scopes: [config.powerBI.scopeBase],
            username: config.powerBI.pbiUsername,
            password: config.powerBI.pbiPassword
        };

        return clientApplication.acquireTokenByUsernamePassword(usernamePasswordRequest);

    };

    // Service Principal auth is the recommended by Microsoft to achieve App Owns Data Power BI embedding
    if (config.powerBI.authenticationMode.toLowerCase() === "serviceprincipal") {
        msalConfig.auth.clientSecret =  config.powerBI.clientSecret
        const clientApplication = new msal.ConfidentialClientApplication(msalConfig);

        const clientCredentialRequest = {
            scopes: [config.powerBI.scopeBase],
        };

        return clientApplication.acquireTokenByClientCredential(clientCredentialRequest);
    }
}

module.exports.getAccessToken = getAccessToken;