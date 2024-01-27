import * as React from 'react';
import { useState, useEffect, useContext } from 'react';
import { ILoadListsProps } from '../models/ILoadListsProps';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as strings from 'MergeWebPartStrings';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import LoadColumns from './LoadColumns';
import RowsDataContext from '../context/RowsDataContext';
import { GraphService } from '../services/GraphService';


const LoadLists: React.FC<ILoadListsProps> = (props: React.PropsWithChildren<ILoadListsProps>) => {

  const [lists, setLists] = useState<IDropdownOption[]>([]);
  const [selectedLists, setSelectedLists] = useState<string[]>([]);
  const [showItems, setShowItems] = useState<boolean>(false);
  const { detailsListData, setDetailsListData } = useContext(RowsDataContext);

  const getAllSpsLists = async () => {
    let client = await props.webContext.msGraphClientFactory.getClient('3');
    let graphService = new GraphService(client);
    const lists = props.siteLists.map(site => {
      return graphService.parseLists(site);
    });
    return lists;
  }

  const getAllLists = async () => {
    try {
      const results = await Promise.all(props.siteLists.map(site => getAllSpsLists()));
      const flattenedAndResolved = [];
      for (const result of results) {
        for (const innerResult of result) {
          const resolvedValue = await innerResult;
          flattenedAndResolved.push(...resolvedValue);
        }
  
        // Clear the state by initializing 'lists' as an empty array
        setLists([]);
        const uniqueItems = [...new Set(flattenedAndResolved.map(item => item.key))].map(key => flattenedAndResolved.find(item => item.key === key));
        setLists((prevLists) => [...prevLists, ...uniqueItems]);
      }
    }catch (error) {
      console.error(error);
    }
  }
  

  useEffect(() => {
    getAllLists();
  }, [props.siteLists]);




  // get selected site and set that to a state. 
  const onListChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item) {
      // Determine the updated selected state of the item
      const updatedSelectedListKeys = item.selected
        ? [...selectedLists, item.key as string]
        : selectedLists.filter(key => key !== item.key);

      // Update the selectedKeys state and then check the length
      setSelectedLists(updatedSelectedListKeys);
      setShowItems(updatedSelectedListKeys.length > 0);
    }
  };


  return (
    <>
      {lists && <Dropdown
        placeholder={strings.SelectList}
        label={strings.SelectList}
        selectedKeys={selectedLists}
        onChange={onListChange}
        multiSelect
        options={lists}
      />}
      {showItems && <LoadColumns webContext={props.webContext} listItems={selectedLists} onChildDataHandler={props.onChildDataHandler} />}
    </>
  );
}

export default LoadLists;