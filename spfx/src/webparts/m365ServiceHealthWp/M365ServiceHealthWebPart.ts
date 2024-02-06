import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption, PropertyPaneDropdown, PropertyPaneLabel, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import PnPTelemetry from "@pnp/telemetry-js";
import * as strings from 'M365ServiceHealthWebPartStrings';

export interface IM365ServiceHealthWebPartProps {
  moreInfo: string;
  moreInfoTooltip: string;
  onlyTiles: boolean;
  showServices: string[];
  tileColumnSize: number;
  tileCompact: boolean;
  tilePageSize: number;
  timeFormat: string;
  timeZone: string;
  title: string;
  webUrl: string;
}

// Reference the solution
import "../../../../dist/m365-service-health.min.js";
declare const M365ServiceHealth: {
  description: string;
  getServices: () => Map<string, string>;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    moreInfo?: string;
    moreInfoTooltip?: string;
    onLoaded?: () => void;
    onlyTiles?: boolean;
    showServices?: string[];
    tileColumnSize?: number;
    tileCompact?: boolean;
    tilePageSize?: number;
    timeFormat?: string;
    timeZone?: string;
    title?: string;
    sourceUrl?: string;
  }) => void;
  moreInfoTooltip: string;
  onlyTiles: boolean;
  tileColumnSize: number;
  tileCompact: boolean;
  tilePageSize: number;
  timeFormat: string;
  timeZone: string;
  title: string;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
  version: string;
};

export default class M365ServiceHealthWebPart extends BaseClientSideWebPart<IM365ServiceHealthWebPartProps> {
  private _hasRendered: boolean = false;
  private _serviceOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    // Opt out of PnP Telemetry
    const telemetry = PnPTelemetry.getInstance();
    telemetry.optOut();

    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the default property values
    if (!this.properties.moreInfoTooltip) { this.properties.moreInfoTooltip = M365ServiceHealth.moreInfoTooltip; }
    if (typeof (this.properties.onlyTiles) === "undefined") { this.properties.onlyTiles = M365ServiceHealth.onlyTiles; }
    if (!this.properties.tileColumnSize) { this.properties.tileColumnSize = M365ServiceHealth.tileColumnSize; }
    if (typeof (this.properties.tileCompact) === "undefined") { this.properties.tileCompact = M365ServiceHealth.tileCompact; }
    if (!this.properties.tilePageSize) { this.properties.tilePageSize = M365ServiceHealth.tilePageSize; }
    if (!this.properties.timeFormat) { this.properties.timeFormat = M365ServiceHealth.timeFormat; }
    if (!this.properties.timeZone) { this.properties.timeZone = M365ServiceHealth.timeZone; }
    if (!this.properties.title) { this.properties.title = M365ServiceHealth.title; }
    if (!this.properties.webUrl) { this.properties.webUrl = this.context.pageContext.web.serverRelativeUrl; }

    // Render the application
    M365ServiceHealth.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type,
      moreInfo: this.properties.moreInfo,
      moreInfoTooltip: this.properties.moreInfoTooltip,
      onlyTiles: this.properties.onlyTiles,
      showServices: this.properties.showServices,
      tileColumnSize: this.properties.tileColumnSize,
      tileCompact: this.properties.tileCompact,
      tilePageSize: this.properties.tilePageSize,
      timeFormat: this.properties.timeFormat,
      timeZone: this.properties.timeZone,
      title: this.properties.title,
      sourceUrl: this.properties.webUrl,
      onLoaded: () => {
        // Clear the service options
        this._serviceOptions = [];

        // Parse the services and set the options
        M365ServiceHealth.getServices().forEach((value, key) => {
          this._serviceOptions.push({
            key: key,
            text: value
          });
        });
      }
    });

    // Set the flag
    this._hasRendered = true;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    M365ServiceHealth.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse(M365ServiceHealth.version);
  }

  protected get disableReactivePropertyChanges(): boolean { return true; }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneSlider('tileColumnSize', {
                  label: strings.TileColumnSizeFieldLabel,
                  max: 6,
                  min: 1,
                  showValue: true
                }),
                PropertyPaneSlider('tilePageSize', {
                  label: strings.TilePageSizeFieldLabel,
                  max: 30,
                  min: 1,
                  showValue: true
                }),
                PropertyPaneToggle('tileCompact', {
                  label: strings.TileCompactFieldLabel,
                  offText: "Standard",
                  onText: "Compact"
                }),
                PropertyPaneToggle('onlyTiles', {
                  label: strings.OnlyTilesFieldLabel,
                  offText: "Full App",
                  onText: "Only Tiles"
                }),
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                PropertyPaneTextField('moreInfo', {
                  label: strings.MoreInfoFieldLabel,
                  description: strings.MoreInfoFieldDescription
                }),
                PropertyPaneTextField('moreInfoTooltip', {
                  label: strings.MoreInfoTooltipFieldLabel,
                  description: strings.MoreInfoTooltipFieldDescription
                }),
                PropertyFieldMultiSelect('showServices', {
                  key: 'showServices',
                  label: strings.ShowServicesFieldLabel,
                  options: this._serviceOptions,
                  selectedKeys: this.properties.showServices
                }),
                PropertyPaneTextField('timeFormat', {
                  label: strings.TimeFormatFieldLabel,
                  description: strings.TimeFormatFieldDescription
                }),
                PropertyPaneDropdown('timeZone', {
                  label: strings.TimeZoneFieldLabel,
                  selectedKey: 'America/New_York',
                  options: [
                    { key: 'America/Anchorage', text: 'America/Anchorage' },
                    { key: 'America/Chicago', text: 'America/Chicago' },
                    { key: 'America/Denver', text: 'America/Denver' },
                    { key: 'America/Los_Angeles', text: 'America/Los Angeles' },
                    { key: 'America/New_York', text: 'America/New York' },
                    { key: 'America/Phoenix', text: 'America/Phoenix' },
                    { key: 'America/Puerto_Rico', text: 'America/Puerto Rico' },
                    { key: 'Pacific/Guam', text: 'Pacific/Guam' },
                    { key: 'Pacific/Honolulu', text: 'Pacific/Honolulu' }
                  ]
                }),
                PropertyPaneLabel('version', {
                  text: "v" + M365ServiceHealth.version
                })
              ]
            }
          ],
          header: {
            description: M365ServiceHealth.description
          }
        }
      ]
    };
  }
}
