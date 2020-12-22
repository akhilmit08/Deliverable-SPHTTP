import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DeliverablesWebPart.module.scss';
import * as strings from 'DeliverablesWebPartStrings';
import {IDeliverableListItem} from '../../models';
import {DeliverableService} from '../../services';

export interface IDeliverablesWebPartProps {
  description: string;
}


export default class DeliverablesWebPart extends BaseClientSideWebPart <IDeliverablesWebPartProps> {

  private deliverableDetailElement: HTMLElement;
  private deliverableService: DeliverableService;

  protected onInit(): Promise<void> {
    this.deliverableService = new DeliverableService(this.context.pageContext.web.absoluteUrl, this.context.spHttpClient);

    return Promise.resolve();
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.deliverables }">
    <div class="${ styles.container }">
      <div class="${ styles.row }">
        <div class="${ styles.column }">
          <span class="${ styles.title }">CCAR Deliverables</span>
  <p class="${ styles.subTitle }">SharePoint HTTP Client</p>
  <button id="getDeliverables" class="${styles.button}">get Delievrables</button>
  <button id="createDeliverable" class="${styles.button}">create Deliverable</button>
  <div id="deliverabledetails"></div>
          </div>
          </div>
          </div>
          </div>`;

          this.deliverableDetailElement = document.getElementById('deliverabledetails');

      document.getElementById('getDeliverables')
        .addEventListener('click', () => {
          this._getDeliverables();
        });
        document.getElementById('createDeliverable')
        .addEventListener('click', () => {
          this._createDeliverable();
        });
  }

  private _getDeliverables(): void {
    this.deliverableService.getDeliverables()
      .then((deliverables: IDeliverableListItem[]) => {
        this._renderDeliverables(this.deliverableDetailElement, deliverables);
      });
  }

  private _createDeliverable(): void {
    const newMission: IDeliverableListItem = <IDeliverableListItem>{
      Title: 'Apollo 18',
      Description1: 'Lovell',
     
    };
    this.deliverableService.createDeliverable(newMission)
      .then(() => {
        this._getDeliverables();
      });

    this._renderDeliverables(this.deliverableDetailElement, null);
  }

  /**
   * Renders collection of missions into the specified HTML element.
   *
   * @private
   * @param {HTMLElement}         element   HTML element to render missions in.
   * @param {IDeliverableListItem[]}  missions  Collection of missions to display.
   * @memberof SharePointHttpClientDemoWebPartWebPart
   */
  private _renderDeliverables(element: HTMLElement, missions: IDeliverableListItem[]): void {
    let missionList: string = '';

    if (missions && missions.length && missions.length > 0) {
      missions.forEach((mission: IDeliverableListItem) => {
        missionList = missionList + `<tr>
        <td>${mission.Title}</td>
      
        <td>${mission.Scenario}</td>
        <td>${mission.Description1}</td>
      </tr>`;
      });
    }

    element.innerHTML = `<table border=1>
      <tr>
        <th>Deliverable</th>
       
        <th>Scenariot</th>
        <th>Description</th>
      </tr>
      <tbody>${missionList}</tbody>
    </table>`;
  }

  protected get dataVersion(): Version {
  return Version.parse('1.0');
}

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              })
            ]
          }
        ]
      }
    ]
  };
}
}
