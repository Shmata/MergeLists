interface Site {
    id: string;
    lists: List[];
}

interface List {
    name: string;
    guid: string;
    items: string[];
}

export class Utilities {
    public isUndefinedNullOrEmpty(value: any): boolean {
        return typeof (value) === "undefined" || value === null || value === "";
    }

    // Helper function to check if an item is in an array
    public containsItem = (array, item) => {
        for (let i = 0; i < array.length; i++) {
        if (array[i] === item) {
            return true;
        }
        }
        return false;
    }

    // This method will return site internal name 
    public getSiteNameFromUrl(url: string | string[] | null): string | null {
        if (typeof url === 'string') {
            if (url.toLowerCase().indexOf('lists') > -1) {
                url = url.toLowerCase().split('/lists')[0];
            }
        } else if (Array.isArray(url)) {
            if (url[0].toLowerCase().indexOf('lists') > -1) {
                url = url[0].toLowerCase().split('/lists')[0];
            } else {
                url = url[0];
            }
        } else {
            return null;
        }

        const parts = url.split('/');
        for (let i = parts.length - 1; i >= 0; i--) {
            const part = parts[i].trim();
            if (part !== '') {
                return part;
            }
        }
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
            const item = parts[0];
            const listTitle = parts[1];
            const siteId = parts[2];
            const listId = parts[3];
            // Create a key by combining site and list name
            const key = `${siteId}-${listId}`;
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

    public convertStringArrayToString(arr: string[]): string {
        // Use the join method with a comma separator
        const concatenatedString = arr.join(',');
        return concatenatedString;
    }


    // Prepare data to register in the Merge list.
    public transformArray(inputArray: string[][]): string[][] {
        const output: string[][] = [];
    
        for (const innerArray of inputArray) {
            const siteIdSet: { [key: string]: { lists: string[]; columns: string[]; listID: Set<string> } } = {};
    
            for (const element of innerArray) {
                const [columnName, listName, siteId, listGuid] = element.split('|');
    
                if (!siteIdSet[siteId]) {
                    siteIdSet[siteId] = { lists: [listName], columns: [columnName], listID: new Set([listGuid]) };
                } else {
                    if (!siteIdSet[siteId].lists.includes(listName)) {
                        siteIdSet[siteId].lists.push(listName);
                    }
                    siteIdSet[siteId].columns.push(columnName);
                    siteIdSet[siteId].listID.add(listGuid);
                }
            }
    
            for (const siteId in siteIdSet) {
                const { lists, columns, listID } = siteIdSet[siteId];
                const listNames = lists.join('|');
                const columnsString = columns.join('|');
                const listIds = Array.from(listID).join('');
                output.push([siteId, listNames, columnsString, listIds]);
            }
        }
    
        return output;
    }
    

    public reverseTransformArray(inputString: string): string[][] {
        const resultArray: string[][] = [];

        const internalArrays = inputString.split(' ');

        for (const concatenatedString of internalArrays) {
            const internalArray = concatenatedString.split('**');
            resultArray.push(internalArray);
        }

        return resultArray;
    }

    public concatenateArraysToSaveInSharePoint(inputArray: string[][]): string {
        const resultArray: string[] = [];

        for (const internalArray of inputArray) {
            const concatenatedString = internalArray.join('**');
            resultArray.push(concatenatedString);
        }

        const finalResult = resultArray.join(' ');
        return finalResult;
    }

    public ensureTaileSlash(url: string): string {
        return (url.endsWith('/')) ? url : url + "/";
    }

    public generateOutputArray(input: string): string[][] {
        const elements = input.split(" ");
        const output: string[][] = [];
      
        elements.forEach((element) => {
          const [siteId, listName, columnsName, listID] = element.split("**");
          const splitedColumnName = columnsName.split("|");
          const resultArray: string[] = [];
      
          splitedColumnName.forEach((columnName) => {
            const result = `${columnName}|${listName}|${siteId}|${listID}`;
            resultArray.push(result);
          });
      
          output.push(resultArray);
        });
      
        return output;
    }
}