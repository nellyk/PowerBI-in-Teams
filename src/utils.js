// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

let config = require(__dirname + "/../config/custom-environment-variables.json");
function getAuthHeader(accessToken) {

    // Function to append Bearer against the Access Token
    return "Bearer ".concat(accessToken);
}

function validateConfig() {
  
    // Validation function to check whether the Configurations are available in the config.powerBI.json file or not

    let guid = require("guid");

    if (!config.powerBI.authenticationMode) {
        return "AuthenticationMode is empty. Please choose MasterUser or ServicePrincipal in config.powerBI.json.";
    }

    if (config.powerBI.authenticationMode.toLowerCase() !== "masteruser" && config.powerBI.authenticationMode.toLowerCase() !== "serviceprincipal") {
        return "AuthenticationMode is wrong. Please choose MasterUser or ServicePrincipal in config.powerBI.json";
    }

    if (!config.powerBI.clientId) {
        return "ClientId is empty. Please register your application as Native app in https://dev.powerbi.com/apps and fill Client Id in config.powerBI.json.";
    }

    if (!guid.isGuid(config.powerBI.clientId)) {
        return "ClientId must be a Guid object. Please register your application as Native app in https://dev.powerbi.com/apps and fill Client Id in config.powerBI.json.";
    }

    if (!config.powerBI.reportId) {
        return "ReportId is empty. Please select a report you own and fill its Id in config.powerBI.json.";
    }

    if (!guid.isGuid(config.powerBI.reportId)) {
        return "ReportId must be a Guid object. Please select a report you own and fill its Id in config.powerBI.json.";
    }

    if (!config.powerBI.workspaceId) {
        return "WorkspaceId is empty. Please select a group you own and fill its Id in config.powerBI.json.";
    }

    if (!guid.isGuid(config.powerBI.workspaceId)) {
        return "WorkspaceId must be a Guid object. Please select a workspace you own and fill its Id in config.powerBI.json.";
    }

    if (!config.powerBI.authorityUrl) {
        return "AuthorityUrl is empty. Please fill valid AuthorityUrl in config.powerBI.json.";
    }

    if (config.powerBI.authenticationMode.toLowerCase() === "masteruser") {
        if (!config.powerBI.pbiUsername || !config.powerBI.pbiUsername.trim()) {
            return "PbiUsername is empty. Please fill Power BI username in config.powerBI.json.";
        }

        if (!config.powerBI.pbiPassword || !config.powerBI.pbiPassword.trim()) {
            return "PbiPassword is empty. Please fill password of Power BI username in config.powerBI.json.";
        }
    } else if (config.powerBI.authenticationMode.toLowerCase() === "serviceprincipal") {
        if (!config.powerBI.clientSecret || !config.powerBI.clientSecret.trim()) {
            return "ClientSecret is empty. Please fill Power BI ServicePrincipal ClientSecret in config.powerBI.json.";
        }

        if (!config.powerBI.tenantId) {
            return "TenantId is empty. Please fill the TenantId in config.powerBI.json.";
        }

        if (!guid.isGuid(config.powerBI.tenantId)) {
            return "TenantId must be a Guid object. Please select a workspace you own and fill its Id in config.powerBI.json.";
        }
    }
}

module.exports = {
    getAuthHeader: getAuthHeader,
    validateConfig: validateConfig,
}