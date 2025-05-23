import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption, PropertyPaneDropdown, PropertyPaneHorizontalRule, PropertyPaneLabel, PropertyPaneLink, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'M365ServiceHealthWebPartStrings';

export interface IM365ServiceHealthWebPartProps {
  darkTheme: boolean;
  listName: string;
  moreInfo: string;
  moreInfoTooltip: string;
  onlyTiles: boolean;
  showServices: string;
  tileColumnSize: number;
  tileCompact: boolean;
  tilePageSize: number;
  timeFormat: string;
  timeZone: string;
  title: string;
  webUrl: string;
}

// Reference the solution
import "main-lib";
import { IBasePropertyPane } from 'gd-sprest-bs/src/propertyPane/types';
declare const M365ServiceHealth: {
  description: string;
  getLogo: () => SVGImageElement;
  getServices: () => Map<string, string>;
  listName: string;
  propertyPaneServices: (key: string, label: string, properties: IM365ServiceHealthWebPartProps) => IBasePropertyPane;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    darkTheme?: boolean;
    displayMode?: DisplayMode;
    envType?: number;
    listName?: string;
    moreInfo?: string;
    moreInfoTooltip?: string;
    onLoaded?: () => void;
    onlyTiles?: boolean;
    showServices?: string;
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
};

export default class M365ServiceHealthWebPart extends BaseClientSideWebPart<IM365ServiceHealthWebPartProps> {
  private _hasRendered: boolean = false;
  private _serviceOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the default property values
    if (!this.properties.listName) { this.properties.listName = M365ServiceHealth.listName; }
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
      darkTheme: this.properties.darkTheme,
      displayMode: this.displayMode,
      envType: Environment.type,
      listName: this.properties.listName,
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

  // Add the logo to the PropertyPane Settings panel
  protected onPropertyPaneRendered(): void {
    const setLogo = setInterval(() => {
      let closeBtn = document.querySelectorAll("div.spPropertyPaneContainer div[aria-label='M365 Service Health property pane'] button[data-automation-id='propertyPaneClose']");
      if (closeBtn) {
        closeBtn.forEach((el: HTMLElement) => {
          let parent = el.parentElement;
          if (parent && !(parent.firstChild as HTMLElement).classList.contains("logo")) { parent.prepend(M365ServiceHealth.getLogo()) }
        });
        clearInterval(setLogo);
      }
    }, 50);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    M365ServiceHealth.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse(this.context.manifest.version);
  }

  protected get disableReactivePropertyChanges(): boolean { return true; }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Basic Settings:",
              groupFields: [
                PropertyPaneToggle('darkTheme', {
                  label: strings.DarkThemeFieldLabel,
                  offText: "Dark theme is currently disabled.",
                  onText: "Dark theme is currently enabled."
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                M365ServiceHealth.propertyPaneServices("showServices", strings.ShowServicesFieldLabel, this.properties)
              ]
            }
          ]
        },
        {
          groups: [
            {
              groupName: "Advanced Settings:",
              groupFields: [
                PropertyPaneTextField('moreInfo', {
                  label: strings.MoreInfoFieldLabel,
                  description: strings.MoreInfoFieldDescription
                }),
                PropertyPaneTextField('moreInfoTooltip', {
                  label: strings.MoreInfoTooltipFieldLabel,
                  description: strings.MoreInfoTooltipFieldDescription
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
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  description: strings.ListNameFieldDescription
                }),
                PropertyPaneTextField('webUrl', {
                  label: strings.WebUrlFieldLabel,
                  description: strings.WebUrlFieldDescription
                })
              ]
            }
          ]
        },
        {
          groups: [
            {
              groupName: "About this app:",
              groupFields: [
                PropertyPaneLabel('version', {
                  text: "Version: " + this.context.manifest.version
                }),
                PropertyPaneLabel('description', {
                  text: M365ServiceHealth.description
                }),
                PropertyPaneLabel('about', {
                  text: "We think adding sprinkles to a donut just makes it better! SharePoint Sprinkles builds apps that are sprinkled on top of SharePoint, making your experience even better. Check out our site below to discover other SharePoint Sprinkles apps, or connect with us on GitHub."
                }),
                PropertyPaneLabel('support', {
                  text: "Are you having a problem or do you have a great idea for this app? Visit our GitHub link below to open an issue and let us know!"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLink('supportLink', {
                  href: "https://www.spsprinkles.com/",
                  text: "SharePoint Sprinkles",
                  target: "_blank"
                }),
                PropertyPaneLink('sourceLink', {
                  href: "https://github.com/spsprinkles/m365-service-health/",
                  text: "View Source on GitHub",
                  target: "_blank"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}