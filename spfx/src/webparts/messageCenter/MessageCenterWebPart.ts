import { Environment, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneHorizontalRule, PropertyPaneLabel,
  PropertyPaneLink, PropertyPaneSlider, PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

//import styles from './MessageCenterWebPart.module.scss';
import * as strings from 'MessageCenterWebPartStrings';

export interface IMessageCenterWebPartProps {
  listName: string;
  moreInfo: string;
  moreInfoTooltip: string;
  tileColumnSize: number;
  tilePageSize: number;
  timeFormat: string;
  timeZone: string;
  title: string;
  webUrl: string;
}

// Reference the solution
import "../../../../dist/message-center.js";
declare const MessageCenter: {
  description: string;
  getLogo: () => SVGImageElement;
  listName: string;
  render: (props: {
    el: HTMLElement;
    context?: any;
    displayMode?: number;
    envType?: number;
    listName?: string;
    moreInfo?: string;
    moreInfoTooltip?: string;
    tileColumnSize?: number;
    tilePageSize?: number;
    timeFormat?: string;
    timeZone?: string;
    title?: string;
    sourceUrl?: string;
  }) => void;
  moreInfoTooltip: string;
  tileColumnSize: number;
  tilePageSize: number;
  timeFormat: string;
  timeZone: string;
  title: string;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
};

export default class MessageCenterWebPart extends BaseClientSideWebPart<IMessageCenterWebPartProps> {

  private _hasRendered: boolean = false;

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the default property values
    if (!this.properties.listName) { this.properties.listName = MessageCenter.listName; }
    if (!this.properties.moreInfoTooltip) { this.properties.moreInfoTooltip = MessageCenter.moreInfoTooltip; }
    if (!this.properties.tileColumnSize) { this.properties.tileColumnSize = MessageCenter.tileColumnSize; }
    if (!this.properties.tilePageSize) { this.properties.tilePageSize = MessageCenter.tilePageSize; }
    if (!this.properties.timeFormat) { this.properties.timeFormat = MessageCenter.timeFormat; }
    if (!this.properties.timeZone) { this.properties.timeZone = MessageCenter.timeZone; }
    if (!this.properties.title) { this.properties.title = MessageCenter.title; }
    if (!this.properties.webUrl) { this.properties.webUrl = this.context.pageContext.web.serverRelativeUrl; }

    // Render the application
    MessageCenter.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type,
      listName: this.properties.listName,
      moreInfo: this.properties.moreInfo,
      moreInfoTooltip: this.properties.moreInfoTooltip,
      tileColumnSize: this.properties.tileColumnSize,
      tilePageSize: this.properties.tilePageSize,
      timeFormat: this.properties.timeFormat,
      timeZone: this.properties.timeZone,
      title: this.properties.title,
      sourceUrl: this.properties.webUrl
    });

    // Set the flag
    this._hasRendered = true;
  }

  //protected onInit(): Promise<void> { }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    MessageCenter.updateTheme(currentTheme);
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Basic Settings:",
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
                  text: MessageCenter.description
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
                  href: "https://github.com/spsprinkles/message-center/",
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
