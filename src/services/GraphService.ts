import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import * as strings from 'MergeWebPartStrings';
import { Utilities } from './Utilities';

import { PermissionKind, hasPermissions } from '../models/IPermissions';
import { PnPClientStorage } from '@pnp/common';
import { TimelinePipe } from '@pnp/core';
import { IWeb } from "@pnp/sp/presets/all";
import { Caching } from "@pnp/queryable";

export class GraphService {

  private _graphClient: MSGraphClientV3;
  private utilities;
  constructor(graphClient: MSGraphClientV3) {
    this._graphClient = graphClient;
    this.utilities = new Utilities();

  }

  public async getAllSpSites(): Promise<IPropertyPaneDropdownOption[]> {
    let allSites = await this._graphClient.api(`/sites?search=*`).get();
    if (!this.utilities.isUndefinedNullOrEmpty(allSites.value)) {
      const siteCollections = allSites.value.map(
        (row) => {
          return {
            text: row.displayName,
            key: row.id,
          };
        }
      );
      return (siteCollections);
    }
  }


  public async parseLists(siteId: string): Promise<IPropertyPaneDropdownOption[]> {
    try {

      let siteLists = await this._graphClient.api(`/sites/${siteId}/lists`).get();
      if (!this.utilities.isUndefinedNullOrEmpty(siteLists.value)) {
        const listsAndLibraries = siteLists.value;
        let siteName = this.utilities.getSiteNameFromUrl(listsAndLibraries[0].webUrl);
        const lists = listsAndLibraries.filter(item => item.list.hidden === false);
        const listOptions: IDropdownOption[] = lists.map(l => {
          return {
            text: `${l.name}|${siteName}`,
            key: `${l.name}|${siteId}|${l.id}`,
          };
        });

        // Add a header item for the site
        const headerItem: IDropdownOption = {
          key: `${listsAndLibraries[0].webUrl}`,
          text: `${strings.ListFor} ${siteName}`,
          itemType: DropdownMenuItemType.Header,
        };

        // Return an array with the header item followed by the lists
        return [headerItem, ...listOptions];
      } else {
        throw new Error(`Failed to retrieve lists and libraries: ${siteLists}`);
      }
    } catch (error) {
      console.error(error);
      throw error;
    }
  }

  public async getAllSpsLists(siteUrl: string): Promise<IDropdownOption[]> {
    try {
      let siteLists = await this._graphClient.api(`/sites/${siteUrl}/lists`).get();

      if (!this.utilities.isUndefinedNullOrEmpty(siteLists.value)) {
        const listsAndLibraries = siteLists.value;
        const lists = listsAndLibraries.filter(item => item.list.hidden === false);
        const listOptions: IDropdownOption[] = lists.map(list => {
          return {
            text: `${list.displayName}`,
            key: `${list.displayName}|${list.id}`,
          };
        });

        return listOptions;
      } else {
        throw new Error(`Failed to retrieve lists and libraries for site URL: ${siteUrl}`);
      }
    } catch (error) {
      console.error(error);
      throw error;
    }
  }

  public async checkCurrentUserPermissions(web: IWeb, listUrl: string, permission: PermissionKind): Promise<boolean> {
    const storage = new PnPClientStorage();
    return storage.local.getOrPut<boolean>(`CurrentUsersHas${permission.toString()}_${listUrl}`, () => {
      return web.getList(listUrl).effectiveBasePermissions.using(this.getDefaultCachingBehavior())().then((result) => {
        return hasPermissions(result, permission);
      }).catch(() => {
        return false;
      });
    }, new Date(new Date().getTime() + 5 * 60000));

  }

  public getDefaultCachingBehavior(): TimelinePipe<any> {
    return Caching({ store: "session", expireFunc: (url: string) => { return new Date(Date.now() + 3600 * 1000); } });
  }

  public async getColumns(siteId: string): Promise<any> {
    let listRequest = siteId.split('|');
    try {
      let columns = await this._graphClient.api(`/sites/${listRequest[1]}/lists/${listRequest[2]}/columns`)
        .get();
      if (!this.utilities.isUndefinedNullOrEmpty(columns.value)) {
        const allColumns = columns.value;
        // Filter the results to include only lists (BaseTemplate 100)
        const cols = allColumns.filter(item => !item.hidden);
        const dropDownColumns: IDropdownOption[] = cols.map(list => {
          return {
            text: `${list.displayName}`,
            key: `${list.name}|${siteId}`,
          };
        });

        // Add a header item for the site
        const headerItem: IDropdownOption = {
          key: '',
          text: `${strings.ColumnFor} ${listRequest[0]}`,
          itemType: DropdownMenuItemType.Header,
        };

        // Return an array with the header item followed by the lists
        return [headerItem, ...dropDownColumns];
      } else {
        throw new Error(`Failed to retrieve columns: ${listRequest[0]}`);
      }
    } catch (error) {
      console.error(error);
      throw error;
    }
  }


  public async getListItems(siteId: string, listGUID: string, selectedColumns: string): Promise<any> {
    let items = await this._graphClient.api(`/sites/${siteId}/lists/${listGUID}/items?expand=fields(select=${selectedColumns})`)
      .get();
    if (!this.utilities.isUndefinedNullOrEmpty(items.value))
      return items.value;
    return null;
  }

  public async getGraphApiCallRequirements(item: string[]): Promise<any> {
    let siteId: string = '';
    let listGUID: string = '';
    let columns: string[] = [];
    let str: string[];
    if (item.length > 0) {
      for (let i of item) {
        str = i.split('|');
        if (str.length > 0) {
          columns.push(str[0]);
          listGUID = str[3];
          siteId = str[2];
        }
      }
    }
    let selectedCols = this.utilities.convertStringArrayToString(columns);
    let res = await this.getListItems(siteId, listGUID, selectedCols);
    return res;
  }

  public async getSecondTestListItems(): Promise<any> {
    let allItems = await this._graphClient.api(`/sites/5jsdrp.sharepoint.com,2c8b17c8-9798-457b-aad0-213b5475cbcb,186a48ff-ee1c-46ba-9e79-6b1f37563892/lists/25D5EA33-3AC6-41BD-9195-CD3AA7FCC100/items?expand=fields(select=Title,Id,Created,Author)`)
      .get();
    return allItems;
  }

}