import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneHorizontalRule, PropertyPaneLabel,
  PropertyPaneLink, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'M365ServiceHealthWebPartStrings';
import { Components, PropertyPane } from "gd-sprest-bs";

export interface IM365ServiceHealthWebPartProps {
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
import "../../../../dist/m365-service-health.min.js";
declare const M365ServiceHealth: {
  description: string;
  getLogo: () => SVGImageElement;
  getServices: () => Map<string, string>;
  listName: string;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    listName?: string;
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
};

export default class M365ServiceHealthWebPart extends BaseClientSideWebPart<IM365ServiceHealthWebPartProps> {
  private _ddlServices: Components.IDropdown;
  private _hasRendered: boolean = false;
  private _serviceItems: Components.IDropdownItem[];

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
      displayMode: this.displayMode,
      envType: Environment.type,
      listName: this.properties.listName,
      moreInfo: this.properties.moreInfo,
      moreInfoTooltip: this.properties.moreInfoTooltip,
      onlyTiles: this.properties.onlyTiles,
      showServices: this.getServices(true),
      tileColumnSize: this.properties.tileColumnSize,
      tileCompact: this.properties.tileCompact,
      tilePageSize: this.properties.tilePageSize,
      timeFormat: this.properties.timeFormat,
      timeZone: this.properties.timeZone,
      title: this.properties.title,
      sourceUrl: this.properties.webUrl,
      onLoaded: () => {
        // Parse the services and set the options
        const services = M365ServiceHealth.getServices();
        this._serviceItems = [];
        services.forEach((value, key) => {
          this._serviceItems.push({
            text: value,
            value: key
          })
        });

        // Set the services
        this.setServices();
      }
    });

    // Set the flag
    this._hasRendered = true;
  }

  // Gets the services to display
  getServices(valuesOnly: boolean = false): string[] | undefined {
    try {
      // Convert the value to an array of dropdown items
      const items = JSON.parse(this.properties.showServices);

      // See if we are returning the values only
      if (valuesOnly) {
        // Parse the items and create an array of the values
        const values: string[] = [];
        for (let i = 0; i < items.length; i++) {
          values.push(items[i].value);
        }

        // Return the values only
        return values;
      }

      // Return items
      return items;
    }
    catch { return undefined; }
  }

  // Sets the services dropdown items and default value
  private setServices(): void {
    // See if the dropdown was already rendered
    if (this._ddlServices) {
      // Update the dropdown properties
      this._ddlServices.setItems(this._serviceItems);

      // See if a value exists
      const currentValue = this.getServices();
      if (currentValue) {
        // Set the selected values
        this._ddlServices.setValue(currentValue);
      }
    }
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

  // Checks if is in debug mode from the query string
  private debug(): boolean {
    // Get the parameters from the query string
    let qs = document.location.search.split('?');
    qs = qs.length > 1 ? qs[1].split('&') : [];
    for (let i = 0; i < qs.length; i++) {
      let qsItem = qs[i].split('=');
      let key = qsItem[0];
      let value = qsItem[1];

      // See if this is the 'debug' key
      if (key === "debug") {
        // Return the item
        return value === "true";
      }
    }
    return false;
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
                PropertyPane.MultiDropdownCheckbox("showServices", {
                  label: strings.ShowServicesFieldLabel,
                  description: "This is a custom control that doesn't affect this app.",
                  placeholder: "Select Services",
                  properties: this.properties,
                  tooltip: "This is my tooltip",
                  items: this._serviceItems,
                  onRendered: (ddl, props) => {
                    // Set the component
                    this._ddlServices = ddl as Components.IDropdown;

                    // See if no items were set, but they exist now
                    if (props?.items === null && this._serviceItems) {
                      // Set the items
                      this._ddlServices.setItems(this._serviceItems);

                      // Set the selected values
                      this.setServices();
                    }
                  }
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
                  description: strings.ListNameFieldDescription,
                  disabled: !this.debug()
                }),
                PropertyPaneTextField('webUrl', {
                  label: strings.WebUrlFieldLabel,
                  description: strings.WebUrlFieldDescription,
                  disabled: !this.debug()
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