/* eslint-disable @typescript-eslint/ban-ts-comment */
import { Injectable } from '@nestjs/common';
import * as msal from '@azure/msal-node';
import { validateConfig } from './utils';
import { EmbedConfig, PowerBiReportDetails } from './models';
import { config } from './consts';

@Injectable()
export class AppService {
  async getEmbedTokenForSingleReportSingleWorkspace(
    reportId,
    datasetIds,
    targetWorkspaceId,
  ) {
    // Add report id in the request
    const formData = {
      reports: [
        {
          id: reportId,
        },
      ],
    };

    // Add dataset ids in the request
    formData['datasets'] = [];
    for (const datasetId of datasetIds) {
      formData['datasets'].push({
        id: datasetId,
      });
    }

    // Add targetWorkspace id in the request
    if (targetWorkspaceId) {
      formData['targetWorkspaces'] = [];
      formData['targetWorkspaces'].push({
        id: targetWorkspaceId,
      });
    }

    const embedTokenApi = 'https://api.powerbi.com/v1.0/myorg/GenerateToken';
    const headers = await this.getRequestHeader();

    // Generate Embed token for single report, workspace, and multiple datasets. Refer https://aka.ms/MultiResourceEmbedToken
    const result = await fetch(embedTokenApi, {
      method: 'POST',
      /* @ts-ignore */
      headers: headers,
      body: JSON.stringify(formData),
    });

    if (!result.ok) throw result;
    return result.json();
  }

  getAuthHeader(accessToken) {
    // Function to append Bearer against the Access Token
    return 'Bearer '.concat(accessToken);
  }

  async getAccessToken() {
    const msalConfig = {
      auth: {
        clientId: config.clientId,
        authority: `${config.authorityUrl}${config.tenantId}`,
        clientSecret: undefined,
      },
    };

    // Check for the MasterUser Authentication
    if (config.authenticationMode.toLowerCase() === 'masteruser') {
      const clientApplication = new msal.PublicClientApplication(msalConfig);

      const usernamePasswordRequest = {
        scopes: [config.scopeBase],
        username: config.pbiUsername,
        password: config.pbiPassword,
      };

      return clientApplication.acquireTokenByUsernamePassword(
        usernamePasswordRequest,
      );
    }

    // Service Principal auth is the recommended by Microsoft to achieve App Owns Data Power BI embedding
    if (config.authenticationMode.toLowerCase() === 'serviceprincipal') {
      msalConfig.auth.clientSecret = config.clientSecret;

      const clientApplication = new msal.ConfidentialClientApplication(
        msalConfig,
      );

      const clientCredentialRequest = {
        scopes: [config.scopeBase],
      };
      console.log('serviceprincipal', clientCredentialRequest);
      const teste = await clientApplication.acquireTokenByClientCredential(
        clientCredentialRequest,
      );

      return teste;
    }
  }

  async getRequestHeader() {
    // Store authentication token
    let tokenResponse;

    // Store the error thrown while getting authentication token
    let errorResponse;

    // Get the response from the authentication request
    try {
      tokenResponse = await this.getAccessToken();
    } catch (err) {
      if (
        err.hasOwnProperty('error_description') &&
        err.hasOwnProperty('error')
      ) {
        errorResponse = err.error_description;
      } else {
        // Invalid PowerBI Username provided
        errorResponse = err.toString();
      }
      return {
        status: 401,
        error: errorResponse,
      };
    }

    // Extract AccessToken from the response
    const token = tokenResponse.accessToken;
    return {
      'Content-Type': 'application/json',
      Authorization: this.getAuthHeader(token),
    };
  }

  async getEmbedParamsForSingleReport(
    workspaceId: string,
    reportId: string,
    additionalDatasetId?: string,
  ) {
    const reportInGroupApi = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/reports/${reportId}`;
    const headers = await this.getRequestHeader();

    console.log('HEADERS', headers);

    // Get report info by calling the PowerBI REST API
    const result = await fetch(reportInGroupApi, {
      method: 'GET',
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-expect-error
      headers: headers,
    });

    if (!result.ok) {
      throw result;
    }

    // Convert result in json to retrieve values
    const resultJson = await result.json();

    // Add report data for embedding
    const reportDetails = new PowerBiReportDetails(
      resultJson.id,
      resultJson.name,
      resultJson.embedUrl,
    );
    const reportEmbedConfig = new EmbedConfig();

    // Create mapping for report and Embed URL
    reportEmbedConfig.reportsDetail = [reportDetails];

    // Create list of datasets
    const datasetIds = [resultJson.datasetId];

    // Append additional dataset to the list to achieve dynamic binding later
    if (additionalDatasetId) {
      datasetIds.push(additionalDatasetId);
    }

    // Get Embed token multiple resources
    reportEmbedConfig.embedToken =
      await this.getEmbedTokenForSingleReportSingleWorkspace(
        reportId,
        datasetIds,
        workspaceId,
      );

    return reportEmbedConfig;
  }

  async getEmbedInfo() {
    // Get the Report Embed details
    try {
      // Get report details and embed token
      const embedParams = await this.getEmbedParamsForSingleReport(
        config.workspaceId,
        config.reportId,
      );

      return {
        accessToken: embedParams.embedToken.token,
        embedUrl: embedParams.reportsDetail,
        expiry: embedParams.embedToken.expiration,
        status: 200,
      };
    } catch (err) {
      console.log('Error', err);
      return {
        status: err.status,
        error: `Error while retrieving report embed details\r\n${
          err.statusText
        }\r\nRequestId: \n${err.headers.get('requestid')}`,
      };
    }
  }

  async getNestleBi() {
    const configCheckResult = validateConfig();

    if (configCheckResult) {
      throw new Error('Error while creating BI');
    }

    try {
      const result = await this.getEmbedInfo();

      return result;
    } catch (error) {
      return error;
    }
  }
}
