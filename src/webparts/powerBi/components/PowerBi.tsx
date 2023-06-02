import * as React from "react";
import styles from "./PowerBi.module.scss";
import { IPowerBiProps, IPowerBiState } from "../../../types";
import { PowerBIEmbed } from "powerbi-client-react";
import { Report, models, VisualDescriptor } from "powerbi-client";
import { IBasicFilter, ISlicer } from "powerbi-models";

export default class PowerBi extends React.Component<
  IPowerBiProps,
  IPowerBiState
> {
  private report: Report;

  constructor(props: IPowerBiProps) {
    super(props);

    this.state = {};
  }

  public render(): React.ReactElement<IPowerBiProps> {
    const { height, loginName, reportId, reportOptions, token, width } =
      this.props;

    const isSlicer = (f: VisualDescriptor) => f.type === "slicer";

    const target: models.IFilterGeneralTarget = {
      column: reportOptions.targetColumn,
      table: reportOptions.targetTable,
    };

    const filter: IBasicFilter = {
      $schema: "http://powerbi.com/product/schema#basic",
      filterType: 1,

      target: {
        column: reportOptions.targetColumn,
        table: reportOptions.targetTable,
      },
      operator: "In",
      values: [loginName],
      requireSingleSelection: false,
    };

    const slicers: Array<ISlicer> = [
      {
        selector: {
          $schema: "http://powerbi.com/product/schema#slicerTargetSelector",
          target: target,
        },
        state: {
          filters: [filter],
        },
      },
    ];

    const embedConfig: models.IReportEmbedConfiguration = {
      type: "report", // Supported types: report, dashboard, tile, visual, qna, paginated report and create
      embedUrl: `https://app.powerbi.com/reportEmbed?reportId=${reportId}`,
      accessToken: token,
      tokenType: models.TokenType.Aad, // Use models.TokenType.Aad for SaaS embed
      settings: {
        panes: {
          filters: {
            expanded: false,
            visible: !reportOptions?.hideFilterPane,
          },
          pageNavigation: {
            visible: !reportOptions?.hidePageNavigation,
          },
        },
        zoomLevel: reportOptions?.zoomLevel,
      },
      slicers: slicers,
    };

    const onRender = () => {
      this.report.getActivePage().then((page) => {
        page.getVisuals().then((visuals) => {
          const allSlicers = visuals.filter(isSlicer);

          allSlicers?.forEach((slicer) => {
            slicer.getSlicerState().then((state) => {
              console.log(`Slicer Config (${slicer.name})`, {
                name: slicer.name,
                targets: state.targets,
                filters: state.filters,
              });
            });
            if (reportOptions?.hideSlicer) {
              slicer.setVisualDisplayState(
                models.VisualContainerDisplayMode.Hidden
              );
            }
          });
        });
      });
    };

    return (
      <div className={styles.powerBi}>
        <style>
          {`.${styles.powerBi} iframe {
            width: ${width}px !important;
            min-height:${height}px !important;
            border: none !important;
          }`}
        </style>
        <div>
          {reportId && token && (
            <PowerBIEmbed
              embedConfig={embedConfig}
              eventHandlers={new Map([["rendered", onRender]])}
              cssClassName={styles.report}
              getEmbeddedComponent={(embeddedReport) => {
                this.report = embeddedReport as Report;
              }}
            />
          )}
          {(!reportId || !token) && (
            <div className={styles.configure}>
              Use the web part property pane to select a workspace and report
              that you want to render.
            </div>
          )}
        </div>
      </div>
    );
  }
}
