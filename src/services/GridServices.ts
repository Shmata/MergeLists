import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { Utilities } from "./Utilities";
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

const ignoredColumns = ["contentType","eTag","odata.type", "odata.id", "odata.etag", "odata.editLink","@odata.etag","fields","fields@odata.context","parentReference"];

export class GridServices {
    private utils: Utilities;
    
    constructor(){
        this.utils = new Utilities();
    }

    // This method is in charge of generate detailsList's columns from api
    public getUniquePropertyNames = (items: any[]) => {
        const uniquePropertyNames = new Set();
        for (const item of items) {
            if (!this.utils.isUndefinedNullOrEmpty(item.fields)) {
                let flds = item.fields;
                for (const propName in flds) {
                    if (Object.prototype.hasOwnProperty.call(flds, propName) && ignoredColumns.indexOf(propName) === -1) {
                        uniquePropertyNames.add(propName);
                    }
                }
            }

        }
        return Array.from(uniquePropertyNames) as any[];
    }

    public removeDuplicates = (cols:any[]) => {
        const uniqueColumnsMap = new Map<string, IColumn>();
        for (const col of cols) {
            uniqueColumnsMap.set(col.key, col);
        }
        return Array.from(uniqueColumnsMap.values());
    }

    public removeFlatArrayDuplicates<T>(arr: T[]): T[] {
        return Array.from(new Set(arr));
    }

    public removeDropDownDuplicates = (cols: IDropdownOption[] ): IDropdownOption[] =>{
        const uniqueColumnsMap = new Map<string, IDropdownOption>();
        for (const col of cols) {
            uniqueColumnsMap.set(col.key as string, col);
        }
        return Array.from(uniqueColumnsMap.values());
    }

    public mapISPColumnsToIColumns = (columns: any[]): any[] => {
        return columns.map((ispColumn) => {
            // Convert minWidth from string to number
            const minWidth = parseInt(ispColumn.minWidth || '100', 10); // Use a default value if it's not present
            return {
                key: ispColumn, // Use a unique key for each column
                name: ispColumn,
                fieldName: ispColumn,
                minWidth: minWidth,
                maxWidth: 180,
                isResizable: true,
            };
        });
    }

    public flattenArrays = (inputArray: string[][]): string[] =>  {
        const mergedArray: string[] = [].concat(...inputArray);
        return mergedArray;
    }

}