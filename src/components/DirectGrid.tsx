import * as React from 'react';
import { IDirectGridProps } from '../models/IDirectGridProps';
import { GraphService } from '../services/GraphService';
import { useEffect, useState } from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import * as strings from 'MergeWebPartStrings';
import { Utilities } from '../services/Utilities';
import { GridServices } from '../services/GridServices';

const DirectGrid: React.FC<IDirectGridProps> = (props: React.PropsWithChildren<IDirectGridProps>) => {

    const utils = new Utilities();
    const gs = new GridServices();

    const [accumulatedResult, setAccumulatedResult] = useState<any[]>([]);
    const [gridColumns, setGridColumns] = useState<any[]>([]);

    const [selectedColumn, setSelectedColumn] = useState<string | undefined>(gridColumns[0]?.text);
    const [filterText, setFilterText] = useState<string>('');

    useEffect(() => {

        const fetchData = async () => {
            if (!utils.isUndefinedNullOrEmpty(props.items)) {
                try {
                    await getItems();
                    //await accumulateData(rows);
                } catch (error) {
                    console.error('Error fetching data:', error);
                }
            }
        };

        fetchData();
    }, [props.items]);

    const fetchColumns = async (items) => {
        setGridColumns((prevColumns) => {
            // Extract unique property names from the new items
            const newColumns = gs.getUniquePropertyNames(items);
            // Combine the new columns with the previous columns, removing duplicates
            const combinedColumns = gs.removeFlatArrayDuplicates([...prevColumns, ...newColumns]);
            return combinedColumns;
        });
    }


    const accumulateData = async (res) => {
        if (res) {
            const unFilteredItems = res.map(i => i.fields);
            const uniqueRes = unFilteredItems.filter(item => !utils.containsItem(accumulatedResult, item));

            if (uniqueRes.length > 0) {
                setAccumulatedResult((prevAccumulatedRes) => [...prevAccumulatedRes, ...uniqueRes]);
            }
        }
    };

    const getItems = async () => {
        const client = await props.webContext.msGraphClientFactory.getClient('3');
        const graphService = new GraphService(client);
        if (props.items.length > 0) {
            for (let i of props.items) {
                graphService.getGraphApiCallRequirements(i).then(async res => {
                    await fetchColumns(res);
                    await accumulateData(res);
                });
            }
        }
    };

    const filteredItems = accumulatedResult.filter((item) => {
        if (!utils.isUndefinedNullOrEmpty(selectedColumn)) {
            const columnValue = item[selectedColumn as string];
            if (columnValue !== undefined && columnValue !== null) {
                return columnValue.toLowerCase().includes(filterText.toLowerCase());
            }
        }
        return true;
    });

    const onFilterTextChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        setFilterText(event.target.value);
    };

    const dropdownOptions: IDropdownOption[] = gridColumns.map((col, i) => ({
        key: i,
        text: col,
    }));

    const onColumnSelectionChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option) {
            setSelectedColumn(option.text as string);
        }
    };

    return (
        <>
            {gridColumns.length > 0 && <TextField label={strings.FilterLabel}
                placeholder={strings.FilterTextboxPlaceholder}
                value={filterText}
                onChange={onFilterTextChange}
            />
            }
            {gridColumns.length > 0 && <Dropdown
                label={strings.SelectColumn}
                selectedKey={selectedColumn}
                options={gs.removeDropDownDuplicates(dropdownOptions)}
                onChange={onColumnSelectionChange}
            />
            }

            {
                gridColumns.length > 0 && accumulatedResult.length > 0 && <DetailsList
                    items={filteredItems}
                    columns={gs.mapISPColumnsToIColumns(gridColumns)}
                    setKey="set"
                    selectionPreservedOnEmptyClick={true}
                    layoutMode={DetailsListLayoutMode.justified}
                />

            }
        </>
    );
}

export default DirectGrid;