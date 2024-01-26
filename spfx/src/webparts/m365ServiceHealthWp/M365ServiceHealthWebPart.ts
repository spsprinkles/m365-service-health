import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'M365ServiceHealthWebPartStrings';

export interface IM365ServiceHealthWebPartProps {
  tileColumnSize: number;
  tilePageSize: number;
  timeFormat: string;
  title: string;
  webUrl: string;
}

// Reference the solution
import "../../../../dist/m365-service-health.min.js";
declare const M365ServiceHealth: {
  description: string;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    tileColumnSize?: number;
    tilePageSize?: number;
    timeFormat?: string;
    title?: string;
    sourceUrl?: string;
  }) => void;
  tileColumnSize: number;
  tilePageSize: number;
  timeFormat: string;
  title: string;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
  version: string;
};

export default class M365ServiceHealthWebPart extends BaseClientSideWebPart<IM365ServiceHealthWebPartProps> {
  private _hasRendered: boolean = false;

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the default property values
    if (!this.properties.tileColumnSize) { this.properties.tileColumnSize = M365ServiceHealth.tileColumnSize; }
    if (!this.properties.tilePageSize) { this.properties.tilePageSize = M365ServiceHealth.tilePageSize; }
    if (!this.properties.timeFormat) { this.properties.timeFormat = M365ServiceHealth.timeFormat; }
    if (!this.properties.title) { this.properties.title = M365ServiceHealth.title; }
    if (!this.properties.webUrl) { this.properties.webUrl = this.context.pageContext.web.serverRelativeUrl; }

    // Render the application
    M365ServiceHealth.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type,
      tileColumnSize: this.properties.tileColumnSize,
      tilePageSize: this.properties.tilePageSize,
      timeFormat: this.properties.timeFormat,
      title: this.properties.title,
      sourceUrl: this.properties.webUrl
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                PropertyPaneTextField('timeFormat', {
                  label: strings.TimeFormatFieldLabel,
                  description: strings.TimeFormatFieldDescription
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
