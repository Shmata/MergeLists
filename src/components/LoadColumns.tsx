import * as React from 'react';
import { useState, useEffect, useContext } from 'react';
import { ILoadColumnsProps } from '../models/ILoadColumnsProps';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as strings from 'MergeWebPartStrings';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DefaultButton } from '@fluentui/react/lib/Button';
import styles from '../styles/Merge.module.scss';
//import { LoadGrid } from './LoadGrid';
import RowsDataContext from '../context/RowsDataContext';
import { GraphService } from '../services/GraphService';
import { Utilities } from '../services/Utilities';
import { SharePointServices } from '../services/SharePointServices';

const LoadColumns: React.FC<ILoadColumnsProps> = (props: React.PropsWithChildren<ILoadColumnsProps>) => {
  const [columns, setColumns] = useState<IDropdownOption[]>([]);
  const [selectedColumns, setSelectedColumns] = useState<string[]>([]);
  const [showButton, setShowButton] = useState<boolean>(false);
  const [isDisable, setIsDisable] = useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false); // handle double click on the Show grid button
  const { detailsListData, setDetailsListData } = useContext(RowsDataContext);
  let utils = new Utilities();

  const getColumnsCall = async () => {
    const client = await props.webContext.msGraphClientFactory.getClient('3');
    const graphService = new GraphService(client);
    const ls = props.listItems.map(async (column) => {
      return await graphService.getColumns(column);
    });
    return ls;
  };

  const getAllColumns = async () => {
    const promises: any = await getColumnsCall();
    // Use Promise.all to wait for all promises to resolve
    Promise.all(promises)
      .then(results => {
        const allFieldOptions = results.reduce((accumulator, fieldOptions) => accumulator.concat(fieldOptions), []);
        if (allFieldOptions.length > 0) {
          //setColumns((prev)=> [...prev, ...allFieldOptions ]);
          setColumns(allFieldOptions);
        }
      })
      .catch(error => {
        console.error(error);
      });
  };

  useEffect(() => {
    if(!utils.isUndefinedNullOrEmpty(props.listItems))
      getAllColumns();
  }, [props.listItems, selectedColumns]);

  // get selected site and set that to a state. 
  const onColumnChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item) {
      const updatedSelectedColumnKeys = item.selected
        ? [...selectedColumns, item.key as string]
        : selectedColumns.filter(key => key !== item.key);

      setSelectedColumns(updatedSelectedColumnKeys);
      setShowButton(updatedSelectedColumnKeys.length > 0);
      setIsDisable(!(updatedSelectedColumnKeys.length > 0));
    }
  };

  const fetchItems = async (columnsArrayData: string[][]) => {
    const client = await props.webContext.msGraphClientFactory.getClient('3');
    const graphService = new GraphService(client);
    if (columnsArrayData.length > 0) {
      for (let i of columnsArrayData) {
        graphService.getGraphApiCallRequirements(i).then(res => {
          setDetailsListData(res);
          props.onChildDataHandler(res);
        });

      }
    }
  };

  const addQueryToSharePoint = async (query: string): Promise<any> => {
    try {
      let services = new SharePointServices(props.webContext);
      await services.addItem( query);

    } catch (error) {
      console.error( error);
    }
  };


  //ShowGrid onClick event handle
  const getSelectedItems = async () => {
    if (!isSubmitting) {
      setIsSubmitting(true);
      try {
        let categorizedColumns = utils.parseColumns(selectedColumns);
        let unripe = utils.transformArray(categorizedColumns);
        let ripeString = utils.concatenateArraysToSaveInSharePoint(unripe);
        await addQueryToSharePoint(ripeString);
        await fetchItems(categorizedColumns);
      } catch (error) {
        console.error(error);
      } finally {
        setIsSubmitting(false);
      }
    }
  };

  return (
    <>
      {columns && <Dropdown
        placeholder={strings.SelectColumn}
        label={strings.SelectColumn}
        selectedKeys={selectedColumns}
        onChange={onColumnChange}
        multiSelect
        options={columns}
      />}
      <p> </p>
      {
        showButton &&
        <DefaultButton
          className={styles.button}
          disabled={isDisable || isSubmitting}
          iconProps={{ iconName: 'AddToShoppingList' }}
          text={strings.ShowListData}
          onClick={getSelectedItems} > </DefaultButton>
      }

    </>
  );
}

export default LoadColumns;