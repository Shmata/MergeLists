import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { setup as pnpSetup } from '@pnp/common';
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";

import { BaseComponentContext } from '@microsoft/sp-component-base';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import * as strings from 'MergeWebPartStrings';
import axios from 'axios';
import { Utilities } from './Utilities';


let sp: SPFI;
const listTitle: string = 'MergeLists';

export class SharePointServices {
    //private spHttpClient: SPHttpClient;
    private utils: Utilities;
    constructor(private context: BaseComponentContext) {
        pnpSetup({
            spfxContext: this.context
        });
        sp = spfi().using(SPFx(this.context));
        //this.spHttpClient = this.context.spHttpClient;
        this.utils = new Utilities();
    }

    public getMergeListInternalName(): string {
        return listTitle;
    }

    public async checkIfListExistAndCreate(): Promise<void> {
        try {
            const result = await sp.web.lists.ensure(listTitle);
            const fields = await result.list.fields();
            const queryFieldExists = fields.some(field => field.InternalName === 'Query');
            if (!queryFieldExists)
                await this.createField();

        } catch (error) {
            console.error('Error ensuring list:', error);
        }

    }

    public async createField(): Promise<void> {
        try {
            const fieldConfig = {
                Title: 'Query',
                FieldTypeKind: 3,
                NumberOfLines: 6,
                RichText: false,
                AppendOnly: false,
                AllowHyperlink: true,
                Group: 'My Group',
            };

            const field = await sp.web.lists.getByTitle(listTitle).fields.add(fieldConfig.Title, fieldConfig.FieldTypeKind, fieldConfig);
            const fieldId = await field.field.select('Id')();
            //console.log(`Field created with ID: ${fieldId.Id}`);
        } catch (error) {
            console.error('Error creating field:', error);
        }
    }

    public getFormDigestValue(): Promise<string> {
        const endpoint = `${this.context.pageContext.site.absoluteUrl}/lists/${listTitle}/_api/contextinfo`;

        return fetch(endpoint, {
            method: "POST",
            headers: { Accept: "application/json;odata=verbose" },
        })
            .then((response) => response.json())
            .then((data) => {
                return data.d.GetContextWebInformation.FormDigestValue;
            });
    }

    public async addItem(queryValue: string): Promise<void> {
        const endpoint = `${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;
        this.getFormDigestValue()
            .then((formDigestValue) => {
                const headers = new Headers();
                headers.append("Accept", "application/json;odata=verbose");
                headers.append("Content-Type", "application/json;odata=verbose");
                headers.append("X-RequestDigest", formDigestValue);

                const requestBody = {
                    __metadata: { type: `SP.Data.${listTitle}ListItem` },
                    Title: `Added by: ${this.context.pageContext.user.displayName}`,
                    Query: queryValue,
                };

                fetch(endpoint, {
                    method: "POST",
                    headers: headers,
                    body: JSON.stringify(requestBody),
                })
                    .then((response) => response.json())
                    .then((data) => {
                        //console.log(data);
                    })
                    .catch((error) => {
                        console.error(error);
                    });
            })
            .catch((error) => {
                console.error(error);
            });

    }

    public async getLastQuery(): Promise<any> {
        try {
            const items = await sp.web.lists.getByTitle(listTitle).items.select("Query", "ID", "Created").orderBy('ID', false).top(1)();
            if (!this.utils.isUndefinedNullOrEmpty(items)) {
                const lastItem = items[0];
                return lastItem.Query;
            } else {
                console.log("No items found in the list.");
                return null;
            }
        } catch (error) {
            console.error(error);
            throw error;
        }
    }

    public async getAllSites(): Promise<IPropertyPaneDropdownOption[]> {
        let firstUrl: string = await this.removeSubSiteFromUrl(this.context.pageContext.web.absoluteUrl, this.context.pageContext.web.serverRelativeUrl);
        let baseUrl = this.ensureTaileSlash(firstUrl);
        try {
            const response = await fetch(
                //
                `${baseUrl}_api/search/query?querytext='contentclass:STS_Site'`,
                {
                    method: "GET",
                    headers: {
                        Accept: "application/json;odata=nometadata",
                        "Content-Type": "application/json",
                    },
                }
            );

            if (response.ok) {
                const data = await response.json();
                const siteCollections = data.PrimaryQueryResult.RelevantResults.Table.Rows.map(
                    (row) => {
                        const properties = row.Cells.reduce((acc, cell) => {
                            acc[cell.Key] = cell.Value;
                            return acc;
                        }, {});
                        return {
                            text: properties.Title,
                            key: properties.Path,
                        };
                    }
                );

                return (siteCollections);
            } else {
                console.error("Error retrieving site collections:", response.statusText);
            }
        } catch (error) {
            console.error("Error retrieving site collections:", error);
        }
    }

    public async removeSubSiteFromUrl(baseUrl: string, pathToRemove: string): Promise<string> {
        const url = new URL(baseUrl);
        const path = url.pathname;

        if (path.startsWith(pathToRemove)) {
            // Remove the sub-site path from the URL
            const newPath = path.substr(pathToRemove.length);
            url.pathname = newPath;
            return url.href;
        } else {
            return baseUrl;
        }
    }

    public async getLists(siteUrl: string): Promise<any> {
        try {
            // Configure your SharePoint URL
            const baseUrl = this.ensureTaileSlash(siteUrl);
            const endpointUrl = `${baseUrl}_api/web/lists`;

            // Define your SharePoint credentials and headers
            const headers = {
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
            };
            let siteName = this.getSiteNameFromUrl(siteUrl);
            // Make a GET request to retrieve lists and libraries
            const response = await axios.get(endpointUrl, { headers });
            if (response.status === 200) {
                const listsAndLibraries = response.data.d.results;
                // Filter the results to include only lists (BaseTemplate 100)
                const lists = listsAndLibraries.filter(item => item.BaseTemplate === 100 || item.BaseTemplate === 107);
                const listOptions: IDropdownOption[] = lists.map(list => {
                    return {
                        text: `${list.Title}|${siteName}`,
                        key: `${list.Title}|${siteName}`,
                    };
                });

                // Add a header item for the site
                const headerItem: IDropdownOption = {
                    key: '',
                    text: `${strings.ListFor} ${siteUrl}`,
                    itemType: DropdownMenuItemType.Header,
                };

                // Return an array with the header item followed by the lists
                return [headerItem, ...listOptions];
            } else {
                throw new Error(`Failed to retrieve lists and libraries: ${response.statusText}`);
            }
        } catch (error) {
            console.error(error);
            throw error;
        }
    }

    public async getColumns(listInternalName: string): Promise<IDropdownOption[]> {
        let listData = listInternalName.split('|');
        let listTitle = listData[0];
        let baseUrl: string = await this.removeSubSiteFromUrl(this.context.pageContext.web.absoluteUrl, this.context.pageContext.web.serverRelativeUrl);
        let subSiteUrl = baseUrl.includes(listData[1]) ? '' : `sites/${listData[1]}`;
        // Make a GET request to the SharePoint REST API
        let endPointServiceUrl = this.ensureTaileSlash(baseUrl + subSiteUrl);
        const response = await fetch(`${endPointServiceUrl}_api/web/lists/getbytitle('${listTitle}')/fields`, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose',
            },
        });
        if (!response.ok) {
            // Handle non-OK response (e.g., 404 Not Found)
            throw new Error(`Request failed with status: ${response.status}`);
        }
        const data = await response.json();
        // Filter the results to include only fields that are not hidden
        const fields = data.d.results.filter(item => !item.Hidden);
        // Build fieldOptions
        const fieldOptions: IDropdownOption[] = fields.map(f => ({
            text: `${f.Title}`,
            key: `${f.EntityPropertyName}:${f.TypeAsString}|${listData[0]}|${listData[1]}`,  //f.InternalName
        }));
        // Add a header item for the site
        const headerItem: IDropdownOption = {
            key: '',
            text: `${strings.ColumnFor} ${listTitle}`,
            itemType: DropdownMenuItemType.Header,
        };
        // Include the header item in fieldOptions
        fieldOptions.unshift(headerItem);
        return fieldOptions;
    }

    // This method will return site internal name 
    private getSiteNameFromUrl(url: string): string | null {
        // Split the URL by '/'
        const parts = url.split('/');
        // Find the last non-empty part
        for (let i = parts.length - 1; i >= 0; i--) {
            const part = parts[i].trim();
            if (part !== '') {
                return part;
            }
        }
        // If no non-empty part is found, return null
        return null;
    }

    // Parse columns, this function will put each columns in a appropriate group.
    public parseColumns(arr: string[]): string[][] {
        const groupedArrays: { [key: string]: string[] } = {};
        // Iterate through the input array
        for (const element of arr) {
            // Split the element into parts using '|'
            const parts = element.split('|');
            // Extract site and list name
            const site = parts[2];
            const listName = parts[1];
            // Create a key by combining site and list name
            const key = `${site}-${listName}`;
            // Add the element to the corresponding group
            if (!groupedArrays[key]) {
                groupedArrays[key] = [];
            }
            groupedArrays[key].push(element);
        }

        // Convert the grouped object into an array of arrays
        const groupedResult: string[][] = [];
        for (const key in groupedArrays) {
            if (Object.prototype.hasOwnProperty.call(groupedArrays, key)) {
                groupedResult.push(groupedArrays[key]);
            }
        }
        return groupedResult;
    }

    public ensureTaileSlash(url: string): string {
        return (url.endsWith('/')) ? url : url + "/";
    }


}