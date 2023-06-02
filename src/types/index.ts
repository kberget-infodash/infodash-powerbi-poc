/**
 * https://docs.microsoft.com/en-us/rest/api/power-bi/groups/getgroups
 * We return the Workspace basic fields needed for the webpart.
 */
export interface PowerBiWorkspace {
  id: string;
  name: string;
}

/**
 * https://docs.microsoft.com/en-us/rest/api/power-bi/groups/getgroups
 * We return the Report basic fields needed for the webpart.
 */
export interface PowerBiReport {
  id: string;
  embedUrl: string;
  name: string;
  webUrl: string;
  datasetId: string;
}

/**
 * Webpart Property Pane fields
 */
export interface IPowerBiWebPartProps {
  // Contains the ID of the selected workspace
  workspaceId: string;
  // Contains the ID of select report
  reportId: string;
  // Defines the min-height of the embedded report
  height: string;
  // Custom report options. More could be added here if you wanted more customization.
  reportOptions: {
    // Hide the filter pane on the report
    hideFilterPane: boolean;
    // Hide the page navigation (tabs)
    hidePageNavigation: boolean;
    // hide the slicers on the page
    hideSlicer: boolean;
    /**
     * On report load the webpart will pass the current user's login to the report
     * You must define the @targetTable and @targetColumn of the slicer you want to pre-filter
     * To see what the table/column values are you can render the report first without filters and check the console for debug output.
     */
    targetColumn: string;
    targetTable: string;

    // Determines the zoom level of the content.
    zoomLevel: number;
  };
}

export interface IPowerBiProps {
  // AAD token for the PowerBi API resource.
  token: string;
  // Contains the ID of select report
  reportId: string;
  /**
   * SPFx API passes in the web part width. This helps when a report is not using a full-width zone
   */
  width: number;
  // Defines the min-height of the embedded report
  height: string;
  // The login name of the current logged in user.
  loginName: string;
  reportOptions: {
    // Hide the filter pane on the report
    hideFilterPane: boolean;
    // Hide the page navigation (tabs)
    hidePageNavigation: boolean;
    // hide the slicers on the page
    hideSlicer: boolean;
    /**
     * On report load the webpart will pass the current user's login to the report
     * You must define the @targetTable and @targetColumn of the slicer you want to pre-filter
     * To see what the table/column values are you can render the report first without filters and check the console for debug output.
     */
    targetColumn: string;
    targetTable: string;

    // Determines the zoom level of the content.
    zoomLevel: number;
  };
}

export interface IPowerBiState {
    // No state used but can be wired up as needed. PowerBI already handled loader
}
