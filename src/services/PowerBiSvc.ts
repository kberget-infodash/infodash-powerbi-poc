import {
  PowerBiWorkspace,
  PowerBiReport,
}
  from "../types";
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

require('powerbi-models');
require('powerbi-client');

export interface IPowerBiSvc {
  GetWorkspaces(context: WebPartContext): Promise<Array<PowerBiWorkspace>>;
  GetReports(context: WebPartContext, workspaceId: string): Promise<Array<PowerBiReport>>;
  GetReport(context: WebPartContext, workspaceId: string, reportId: string): Promise<PowerBiReport>;
}

export class PowerBiSvc implements IPowerBiSvc {

  // Static PowerBI resource URI
  private static powerbiApiResourceId = "https://analysis.windows.net/powerbi/api";
  // Static powerBI groups (workspaces) endpoint
  private static workspacesUrl = "https://api.powerbi.com/v1.0/myorg/groups/";
  

  /**
   * 
   * @param context The current webpart's context. Used for the built-in AAD http client for making authenticated calls
   * @returns A list of the workspaces the user has access to.
   */
  public async GetWorkspaces(context: WebPartContext): Promise<Array<PowerBiWorkspace>> {
    try {
      let reqHeaders: HeadersInit = new Headers();
      reqHeaders.append("Accept", "*");

      const pbiClient: AadHttpClient = await context.aadHttpClientFactory.getClient(PowerBiSvc.powerbiApiResourceId);
      const response: HttpClientResponse = await pbiClient.get(PowerBiSvc.workspacesUrl, AadHttpClient.configurations.v1, { headers: reqHeaders });
      const jsonResponse: { value: PowerBiWorkspace[] } = await response.json();

      return Promise.resolve(jsonResponse.value);
    } catch (error) {
      return Promise.reject(error);
    }
  }

  /**
   * 
   * @param context The current webpart's context. Used for the built-in AAD http client for making authenticated calls
   * @param workspaceId The workspace that you want to query for a reports
   * @returns A list of reports for the designated workspace
   */
  public async GetReports(context: WebPartContext, workspaceId: string): Promise<Array<PowerBiReport>> {

    try {
      let reportsUrl = PowerBiSvc.workspacesUrl + workspaceId + "/reports/";
      let reqHeaders: HeadersInit = new Headers();
      reqHeaders.append("Accept", "*");

      const pbiClient: AadHttpClient = await context.aadHttpClientFactory.getClient(PowerBiSvc.powerbiApiResourceId);
      const response: HttpClientResponse = await pbiClient.get(reportsUrl, AadHttpClient.configurations.v1, { headers: reqHeaders });
      const jsonResponse: { value: PowerBiWorkspace[] } = await response.json();
      const reports: Array<PowerBiReport> = jsonResponse.value.map((report: PowerBiReport) => {
        return {
          id: report.id,
          embedUrl: report.embedUrl,
          name: report.name,
          webUrl: report.webUrl,
          datasetId: report.datasetId,
        };
      });

      return Promise.resolve(reports);
    } catch (error) {
      return Promise.reject(error);
    }

  }

  /**
   * 
   * @param context The current webpart's context. Used for the built-in AAD http client for making authenticated calls
   * @param workspaceId The workspace that you want to query for a report
   * @param reportId Provide the report ID of the report you want to fetch
   * @returns Report Profile for embedding. We only use the id
   */
  public async GetReport(context: WebPartContext, workspaceId: string, reportId: string): Promise<PowerBiReport> {

    try {

      let reportUrl = PowerBiSvc.workspacesUrl + workspaceId + "/reports/" + reportId + "/";
      let reqHeaders: HeadersInit = new Headers();
      reqHeaders.append("Accept", "*");

      const pbiClient: AadHttpClient = await context.aadHttpClientFactory.getClient(PowerBiSvc.powerbiApiResourceId);
      const response: HttpClientResponse = await pbiClient.get(reportUrl, AadHttpClient.configurations.v1, { headers: reqHeaders });
      const jsonResponse: PowerBiReport = await response.json();
      const report: PowerBiReport = {
        id: jsonResponse.id,
        embedUrl: jsonResponse.embedUrl,
        name: jsonResponse.name,
        webUrl: jsonResponse.webUrl,
        datasetId: jsonResponse.datasetId,
      };

      return Promise.resolve(report);
    } catch (error) {
      return Promise.reject(error);
    }
  }
}