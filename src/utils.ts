import { config } from './consts';

export function validateConfig() {
  // Validation function to check whether the Configurations are available in the config.js file or not

  if (!config.authenticationMode) {
    return 'AuthenticationMode is empty. Please choose MasterUser or ServicePrincipal in config.js.';
  }

  if (
    config.authenticationMode.toLowerCase() !== 'masteruser' &&
    config.authenticationMode.toLowerCase() !== 'serviceprincipal'
  ) {
    return 'AuthenticationMode is wrong. Please choose MasterUser or ServicePrincipal in config.js';
  }

  if (!config.clientId) {
    return 'ClientId is empty. Please register your application as Native app in https://dev.powerbi.com/apps and fill Client Id in config.js.';
  }

  if (!config.reportId) {
    return 'ReportId is empty. Please select a report you own and fill its Id in config.js.';
  }

  if (!config.workspaceId) {
    return 'WorkspaceId is empty. Please select a group you own and fill its Id in config.js.';
  }

  if (!config.authorityUrl) {
    return 'AuthorityUrl is empty. Please fill valid AuthorityUrl in config.js.';
  }

  if (config.authenticationMode.toLowerCase() === 'masteruser') {
    if (!config.pbiUsername || !config.pbiUsername.trim()) {
      return 'PbiUsername is empty. Please fill Power BI username in config.js.';
    }

    if (!config.pbiPassword || !config.pbiPassword.trim()) {
      return 'PbiPassword is empty. Please fill password of Power BI username in config.js.';
    }
  } else if (config.authenticationMode.toLowerCase() === 'serviceprincipal') {
    if (!config.clientSecret || !config.clientSecret.trim()) {
      return 'ClientSecret is empty. Please fill Power BI ServicePrincipal ClientSecret in config.js.';
    }

    if (!config.tenantId) {
      return 'TenantId is empty. Please fill the TenantId in config.js.';
    }
  }
}
