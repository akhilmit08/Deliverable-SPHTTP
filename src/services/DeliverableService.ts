import {
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions
  } from '@microsoft/sp-http';

import { IDeliverableListItem } from '../models';

const LIST_API_ENDPOINT: string = `/_api/web/lists/getbytitle('Deliverables')`;
const SELECT_QUERY: string = '$select=Id,Title,Scenario,Status,Description1,LOB/Id,LOB/Title&$expand=LOB/Id';
export class DeliverableService {


    private _spHttpOptions: any = {
        getNoMetadata: <ISPHttpClientOptions>{
          headers: { 'ACCEPT': 'application/json; odata.metadata=none' }
        },
        getFullMetadata: <ISPHttpClientOptions>{
          headers: { 'ACCEPT': 'application/json; odata.metadata=full' }
        },
        postNoMetadata: <ISPHttpClientOptions>{
          headers: {
            'ACCEPT': 'application/json; odata.metadata=none',
            'CONTENT-TYPE': 'application/json',
          }
        },
        updateNoMetadata: <ISPHttpClientOptions>{
          headers: {
            'ACCEPT': 'application/json; odata.metadata=none',
            'CONTENT-TYPE': 'application/json',
            'X-HTTP-Method': 'MERGE'
          }
        },
        deleteNoMetadata: <ISPHttpClientOptions>{
          headers: {
            'ACCEPT': 'application/json; odata.metadata=none',
            'CONTENT-TYPE': 'application/json',
            'X-HTTP-Method': 'DELETE'
          }
        }
      };
      constructor(private siteAbsoluteUrl: string, private client: SPHttpClient) { }

       /**
   * Return collection of all NASA Apollo missions.
   *
   * @returns {IMission[]}      Collection of missions.
   * @memberof MissionService
   */
  public getDeliverables(): Promise<IDeliverableListItem[]> {
    let promise: Promise<IDeliverableListItem[]> = new Promise<IDeliverableListItem[]>((resolve, reject) => {
      this.client.get(`${this.siteAbsoluteUrl}${LIST_API_ENDPOINT}/items?${SELECT_QUERY}`,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getNoMetadata
      ) // get response & parse body as JSON
        .then((response: SPHttpClientResponse): Promise<{ value: IDeliverableListItem[] }> => {
          return response.json();
        }) // get parsed response as array, and return
        .then((response: { value: IDeliverableListItem[] }) => {
          resolve(response.value);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  /**
   * Retrieve the entity type as a string for the list
   *
   * @private
   * @returns {Promise<string>}
   * @memberof MissionService
   */
  private _getItemEntityType(): Promise<string> {
    let promise: Promise<string> = new Promise<string>((resolve, reject) => {
      this.client.get(`${this.siteAbsoluteUrl}${LIST_API_ENDPOINT}?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getNoMetadata
      )
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          resolve(response.ListItemEntityTypeFullName);
        })
        .catch((error: any) => {
          reject(error);
        });
    });
    return promise;
  }

  public createDeliverable(newMission: IDeliverableListItem): Promise<void> {
    let promise: Promise<void> = new Promise<void>((resolve, reject) => {
      // first, get the type of thing we're creating...
      this._getItemEntityType()
        .then((spEntityType: string) => {
          // create item to create
          let newListItem: IDeliverableListItem = newMission;
          // add SP-required metadata
          newListItem['@odata.type'] = spEntityType;

          // build request
          let requestDetails: any = this._spHttpOptions.postNoMetadata;
          requestDetails.body = JSON.stringify(newListItem);

          // create the item
          return this.client.post(`${this.siteAbsoluteUrl}${LIST_API_ENDPOINT}/items`,
            SPHttpClient.configurations.v1,
            requestDetails
          );
        })
        .then((response: SPHttpClientResponse): Promise<IDeliverableListItem> => {
          return response.json();
        })
        .then((newSpListItem: IDeliverableListItem): void => {
          resolve();
        })
        .catch((error: any) => {
          reject(error);
        });
    });
    return promise;
  }


}