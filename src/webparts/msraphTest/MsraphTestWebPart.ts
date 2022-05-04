import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MsraphTestWebPart.module.scss';
import * as strings from 'MsraphTestWebPartStrings';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IMsraphTestWebPartProps {
  description: string;
}

export default class MsraphTestWebPart extends BaseClientSideWebPart<IMsraphTestWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.msraphTest} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ''
    }">
      <div class="${styles.welcome}">
        <img alt="" src="${
          this._isDarkTheme
            ? require('./assets/welcome-dark.png')
            : require('./assets/welcome-light.png')
        }" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(
          this.context.pageContext.user.displayName
        )}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(
          this.properties.description
        )}</strong></div>
      </div>
      <div>
        <h3>Welcome to the Microsoft Graph Client Demo!</h3>
      </div>
      <div class="calendarEvents"></div>
    </section>`;
    // getting the 1st element from the list of calendar events
    const calendarEventsDomElement: HTMLElement =
      this.domElement.getElementsByClassName(
        'calendarEvents'
      )[0] as HTMLElement;

    // creating a spinner while fetching the data
    this.context.statusRenderer.displayLoadingIndicator(
      calendarEventsDomElement,
      'Loading your events from your calendar'
    );
    // fetching the calendar events from the graph, removing the spinner and rendering the list of calendar events
    this.getCalendarEvents().then((calendarEvents: MicrosoftGraph.Event[]) => {
      this.context.statusRenderer.clearLoadingIndicator(
        calendarEventsDomElement
      );
      //   rendering the html elements and calendar events?
      this._renderCalendarEvents(calendarEventsDomElement, calendarEvents);
    });
  }

  public getCalendarEvents(): Promise<MicrosoftGraph.Event[]> {
    return new Promise<MicrosoftGraph.Event[]>((resolve, reject) => {
      this.context.msGraphClientFactory
        .getClient()
        // promise
        .then((msGraphClient: MSGraphClient) => {
          // calendar events is any
          msGraphClient
            .api('/me/calendar/events?$top=9')
            .get((err: any, calendarEvents: any, rawResponse: any) => {
              //   fulfilling the promise
              // the value is the array of hte properties on the graph, need to check the graph explorer to get that info.
              resolve(calendarEvents.value);
            });
          // graph only works on live environments, can't do it on the workbench, need to deploy ot the app catalog
          // use gulp build && gulp bundle --ship && gulp package-solution --ship to deploy
        });
    });
  }

  private _renderCalendarEvents(
    element: HTMLElement,
    events: MicrosoftGraph.Event[]
  ): void {
    let htmlTableRows: string = '';

    // id is a very long number so truncating it after 5 characters
    if (events && events.length && events.length > 0) {
      events.forEach((event: MicrosoftGraph.Event) => {
        htmlTableRows =
          htmlTableRows +
          `<tr>
              <td>${event.id.substr(0, 5)}...</td> 
              <td>${event.start.dateTime} (${event.start.timeZone})</td>
              <td>${event.end.dateTime} (${event.end.timeZone})</td>
              <td>${event.subject}</td>
              </tr>`;
      });
    }
    element.innerHTML = `
    <table border=1>
    <tr>
    <th>ID</th>
    <th>Start</th>
    <th>End</th>
    <th>Subject</th>
    </tr>
    <tbody>${htmlTableRows}</tbody>
    </table>
    `;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty(
      '--linkHovered',
      semanticColors.linkHovered
    );
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
