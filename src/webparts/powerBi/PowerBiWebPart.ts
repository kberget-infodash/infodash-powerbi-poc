import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  AadTokenProviderFactory,
  AadTokenProvider,
  HttpClient,
} from "@microsoft/sp-http";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import PowerBi  from "./components/PowerBi";
import { IPowerBiProps } from "../../types";
import { PowerBiSvc, IPowerBiSvc } from "../../services/PowerBiSvc";
import { PowerBiReport, PowerBiWorkspace, IPowerBiWebPartProps} from "../../types";

export default class PowerBiWebPart extends BaseClientSideWebPart<IPowerBiWebPartProps> {

  // Custom service that queries power bi for data
  private powerBiSvc: IPowerBiSvc = new PowerBiSvc();
  // Built-in token provider for generating auth tokens on behalf of the user
  private aadTokenProviderFactory: AadTokenProviderFactory;
  // Contains a list of workspaces (after successsfull call to powerbi svc)
  private workspaceOptions: IDropdownOption[];
  // Switch when workspaces have been fetched successfully.
  private workspacesFetched: boolean = false;
  // Contains a list of report (after successsful call to powerbi svc)
  private reportOptions: IDropdownOption[];
  // Switch when reports have been fetched successfully.
  private reportsFetched: boolean = false;
  // Contains the current user's PowerBi token.
  // TODO: Caching?
  private token: string;

  /**
   * 
   * @returns Return a list of the workspaces that the user has access to
   */
  private async fetchWorkspaceOptions(): Promise<IDropdownOption[]> {
    const workspaces: Array<PowerBiWorkspace> =
      await this.powerBiSvc.GetWorkspaces(this.context);
    const options: Array<IDropdownOption> = new Array<IDropdownOption>();
    workspaces.map((workspace: PowerBiWorkspace) => {
      options.push({ key: workspace.id, text: workspace.name });
    });

    return Promise.resolve(options);
  }

  /**
   * 
   * @returns Return a list of the reports that the user has access to
   */
  private async fetchReportOptions(): Promise<IDropdownOption[]> {
    const workspaces: Array<PowerBiWorkspace> =
      await this.powerBiSvc.GetReports(
        this.context,
        this.properties.workspaceId
      );
    const options: Array<IDropdownOption> = new Array<IDropdownOption>();
    workspaces.map((report: PowerBiReport) => {
      options.push({ key: report.id, text: report.name });
    });

    return Promise.resolve(options);
  }

  public async onInit(): Promise<void> {
    await super.onInit();

    // Instantiate the token factory service scope
    this.aadTokenProviderFactory = this.context.serviceScope.consume(
      AadTokenProviderFactory.serviceKey
    );

    // Initialize the token provider
    const tokenProvider: AadTokenProvider =
      await this.aadTokenProviderFactory.getTokenProvider();

    /**
     * Fetch a toke for the user.
     * Refer to read me to ensure you have granted the proper app permissions to the 
     * SharePoint Web Application Service Principal in AAD
     */
    this.token = await tokenProvider.getToken(
      "https://analysis.windows.net/powerbi/api"
    );
  }

  public async render(): Promise<void> {
    const element: React.ReactElement<IPowerBiProps> = React.createElement(
      PowerBi,
      {
        height: this.properties.height,
        loginName: this.context.pageContext.user.loginName,
        reportId: this.properties.reportId,
        reportOptions: this.properties.reportOptions,
        token: this.token,
        width: this.width,
      }
    );

    ReactDom.render(element, this.domElement);

    return Promise.resolve();
  }

  /**
   * 
   * @returns Start to fetch workspace and reports when property pane has been opened
   */
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if (this.workspacesFetched && this.reportsFetched) {
      return;
    }

    if (this.properties.workspaceId && !this.reportsFetched) {
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Calling Power BI Service API to get reports"
      );
      const reportOptions: Array<IDropdownOption> =
        await this.fetchReportOptions();
      this.reportOptions = reportOptions;
      this.reportsFetched = true;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    }

    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "Calling Power BI Service API to get workspaces"
    );
    const workspaceOptions: Array<IDropdownOption> =
      await this.fetchWorkspaceOptions();
    this.workspaceOptions = workspaceOptions;
    this.workspacesFetched = true;
    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();

    return Promise.resolve();
  }

  /**
   * 
   * @param propertyPath The property that was changed. JSON is flattened so reportOptions will reportOptions.<property>
   * @param oldValue The previous value of the property pane field
   * @param newValue The new value of the property pane field
   * @returns 
   */
  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === "workspaceId" && newValue) {
      // reset report settings
      this.properties.reportId = "";
      this.reportOptions = [];
      this.reportsFetched = false;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Calling Power BI Service API to get reports"
      );
      const reportOptions: Array<IDropdownOption> =
        await this.fetchReportOptions();
      this.reportOptions = reportOptions;
      this.reportsFetched = true;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    }


    if (propertyPath === "reportId" && newValue) {
      this.render();
    }

    if ((propertyPath == 'reportOptions.targetColumn' || propertyPath === 'reportOptions.targetColumn') && this.properties?.reportOptions?.targetColumn && this.properties?.reportOptions?.targetTable) {
      this.render();
    }

    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * When the page has been resized, kick off a render to ensure the report knows 
   * the new avaialble width
   */
  protected onAfterResize(newWidth: number): void {
    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              isCollapsed: false,
              groupName: "PowerBI Report Connection",
              groupFields: [
                PropertyPaneDropdown("workspaceId", {
                  label: "Select a Workspace",
                  options: this.workspaceOptions,
                  disabled: !this.workspacesFetched,
                }),
                PropertyPaneDropdown("reportId", {
                  label: "Select a Report",
                  options: this.reportOptions,
                  disabled: !this.reportsFetched,
                }),
              ],
            },
            {
              isCollapsed: true,
              groupName: "Slicer Configuration",
              groupFields: [
                PropertyPaneTextField("reportOptions.targetTable", {
                  label: "Target Table",
                  description: `Define the target table name in the slicer.`,
                }),
                PropertyPaneTextField("reportOptions.targetColumn", {
                  label: "Target Column",
                  description: `Define the target column name in the slicer.`,
                }),
                PropertyPaneToggle("reportOptions.hideSlicer", {
                  label: "Hide Slicers on Load",
                  onText: "Yes",
                  offText: "No",
                }),
              ]
            },
            {
              isCollapsed: true,
              groupName: "Report Options",
              groupFields: [
                PropertyPaneTextField("height", {
                  label: "Canvas Height (px)",
                  description: 'You must set a height for the report. The report will always use the available width.'
                }),
                PropertyPaneToggle("reportOptions.hideFilterPane", {
                  label: "Hide Filter Pane",
                  onText: "Yes",
                  offText: "No",
                }),
                PropertyPaneToggle("reportOptions.hidePageNavigation", {
                  label: "Hide Page Navigation",
                  onText: "Yes",
                  offText: "No",
                }),
                PropertyPaneSlider("reportOptions.zoomLevel", {
                  label: "Zoom Level",
                  max: 2,
                  min: 0.5,
                  step: 0.1,
                }),
              ]
            }
          ],
          displayGroupsAsAccordion: true
        },
      ],
    };
  }
}
