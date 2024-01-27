import * as React from 'react';
import { useState, useEffect,useContext } from 'react';
import { IShowPanelProps } from "../models/IShowPanelProps";
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import { ChoiceGroup, IChoiceGroupOption, Separator, Text } from '@fluentui/react';
import * as strings from 'MergeWebPartStrings';
import styles from '../styles/Merge.module.scss';
import { GraphService } from '../services/GraphService';

import { Dropdown,  IDropdownOption } from '@fluentui/react/lib/Dropdown';
import LoadLists from './LoadLists';
import RowsDataContext from '../context/RowsDataContext';

const ShowPanel: React.FC<IShowPanelProps> = (props: React.PropsWithChildren<IShowPanelProps>) =>{

    const { isPanelOpen, onPanelDismiss } = props;
    const [visibilityOption, setVisibilityOption] = useState<string>("");

    const [selectedKeys, setSelectedKeys] = useState<string[]>([]);
    const [dropdownItems, setDropdownItems] = useState<IDropdownOption[]>([]); 
    const [showLists, setShowLists] = useState<boolean>(false);
    const { detailsListData, setDetailsListData } = useContext(RowsDataContext);

    const loadAllSitesGraphApi = async ()=>{
      let client = await props.webContext.msGraphClientFactory.getClient('3');
      let graphService = new GraphService(client);
      graphService.getAllSpSites().then( s =>{
          setDropdownItems(s);
      });

    }

    const configurationAdminOptions: IChoiceGroupOption[] = [{
        key: 'USERS',
        text: strings.VisibilityOptionUsers
      },
      {
        key: 'ADMINISTRATORS',
        text: strings.VisibilityOptionAdministrators
      },
      {
        key: 'HIDE',
        text: strings.VisibilityOptionHide
      }];

      // get selected site and set that to a state. 
      const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        if (item) {
          // Determine the updated selected state of the item
          const updatedSelectedKeys = item.selected
            ? [...selectedKeys, item.key as string]
            : selectedKeys.filter(key => key !== item.key);
      
          // Update the selectedKeys state and then check the length
          setSelectedKeys(updatedSelectedKeys);
          setShowLists(updatedSelectedKeys.length > 0);
        }
      };  

    let style: React.CSSProperties = {
        fontSize: `0px`
    };
   
    useEffect(() => {
        loadAllSitesGraphApi(); 
        
      }, [isPanelOpen, selectedKeys])
    
    

    return (
        <>
      <Panel
        isOpen={isPanelOpen}
        onDismiss={onPanelDismiss}
        headerText={strings.MergeLists}
        closeButtonAriaLabel={strings.CloseButtonText}
        type={PanelType.medium}
      >
        {dropdownItems && <Dropdown
          placeholder={strings.SelectSite}
          label={strings.SelectSite}
          selectedKeys={selectedKeys}
          onChange={onChange}
          multiSelect
          options={dropdownItems}
        />}
        {
          showLists && <LoadLists webContext={props.webContext} siteLists={selectedKeys} onChildDataHandler={props.onChildDataHandler}/>
        }
      </Panel>
    </>
    );
}

export default ShowPanel;