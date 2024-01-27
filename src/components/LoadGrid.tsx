import * as React from 'react';
import { ILoadGridProps } from '../models/ILoadGridProps';
import { useEffect, useState } from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import * as strings from 'MergeWebPartStrings';
import { Utilities } from '../services/Utilities';
import { GridServices } from '../services/GridServices';

const LoadGrid: React.FC<ILoadGridProps> = (props: React.PropsWithChildren<ILoadGridProps>) => {
    const [gridColumns, setGridColumns] = useState<any[]>();
    const [cols, setCols] = useState<IColumn[]>([]);
    const [filterText, setFilterText] = useState<string>('');
    const [selectedColumn, setSelectedColumn] = useState<string | undefined>(cols[0]?.key);
    const utils = new Utilities();
    const gs = new GridServices();

    useEffect(() => {
        setGridColumns(gs.getUniquePropertyNames(props.items));
    }, [props.items]);

    useEffect(() => {
        if (gridColumns) {
            setCols((prev) => {
                const newColumns = gs.mapISPColumnsToIColumns(gridColumns);
                // Use a Set to remove duplicates
                const uniqueColumns = new Set([...prev, ...newColumns]);
                return Array.from(uniqueColumns);
            });
        }
    }, [gridColumns]);


    const filteredItems = props.items.filter((item) => {
        
        if (!utils.isUndefinedNullOrEmpty(selectedColumn)) {
            // Find the corresponding property in the item object using the selectedColumn key
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

    const dropdownOptions: IDropdownOption[] = cols.map((col) => ({
        key: col.key,
        text: col.name,
    }));

    const onColumnSelectionChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option) {
          setSelectedColumn(option.key as string);
        }
    };

    
    return (
        <>
            <TextField label={strings.FilterLabel}
                placeholder={strings.FilterTextboxPlaceholder}
                value={filterText}
                onChange={onFilterTextChange}
              />
            {
                cols && <Dropdown
                    label={strings.SelectColumn}
                    selectedKey={selectedColumn}
                    options={gs.removeDropDownDuplicates(dropdownOptions)}
                    onChange={onColumnSelectionChange}
                />
            }
            {
                cols && <DetailsList
                items={filteredItems}
                columns={gs.removeDuplicates(cols)}
                setKey="set"
                selectionPreservedOnEmptyClick={true}
                layoutMode={DetailsListLayoutMode.justified}
            />
            }
        </>
    );
}

export default LoadGrid;