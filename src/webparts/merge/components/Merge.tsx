import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './Merge.module.scss';
import { IMergeProps } from '../../../models/IMergeProps';
import ShowPanel from '../../../components/ShowPanel';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import RowsDataContext from '../../../context/RowsDataContext';
import LoadGrid from '../../../components/LoadGrid';
//import { SharePointServices } from '../../../services/SharePointServices';
//import LoadColumns from '../../../components/LoadColumns';
import { PermissionKind, hasPermissions } from '../../../models/IPermissions';
import { Utilities } from '../../../services/Utilities';
import DirectGrid from '../../../components/DirectGrid';

const Merge: React.FC<IMergeProps> = (props: React.PropsWithChildren<IMergeProps>) => {
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [whoCanSee, setWhoCanSee] = useState<boolean>(false);
  const [canEdit, setCanEdit] = useState<boolean>(false);
  const [detailsListData, setDetailsListData] = useState<any>([]);
  const [callLoadColumns,setCallLoadColumns] = useState<boolean>(false);
  const [queryData, setQueryData] = useState<any>([]);
  const [columns, setColumns] = useState([]);
  const [accumulatedResult, setAccumulatedResult] = useState<any[]>([]);

  const utils = new Utilities();
  let style: React.CSSProperties = {
    fontSize: `${props.buttonSize}px`
  };

  if (props.themeVariant) {
    const { semanticColors }: IReadonlyTheme = props.themeVariant;
    style = {
      fontSize: `${props.buttonSize}px`,
      backgroundColor: semanticColors?.primaryButtonBackground,
      color: semanticColors?.primaryButtonText
    };
  }


  const onShowPanel = () => {
    setIsPanelOpen(!isPanelOpen);
  };

  const handlePanelDismiss = () => {
    setIsPanelOpen(false);
  };

  const checkSiteAdmin = () =>{
    let apiUrl = `${props.webContext.pageContext.web.absoluteUrl}/_api/web/EffectiveBasePermissions`;
    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
      },
    }).then(response => response.json()).then(data => {
      const hasViewListItemsPermission = hasPermissions(data.d.EffectiveBasePermissions, PermissionKind.ManageWeb);
      setWhoCanSee(hasViewListItemsPermission);
      setCanEdit(hasViewListItemsPermission);
    }).catch(error => {
      console.log(error);
    });

  }

  const accumulateData = async (res) => {
    if (res) {
      const unFilteredItems = res.map(i => i.fields);
      // Check for duplicates and accumulate only unique items
      const uniqueRes = unFilteredItems.filter(item => !utils.containsItem(accumulatedResult, item));
      
      if (uniqueRes.length > 0) {
        // Accumulate uniqueRes.fields into accumulatedResult
        setAccumulatedResult((prevAccumulatedRes) => [...prevAccumulatedRes, ...uniqueRes]);
      }
    }
  };
  
  const handleDataFromLoadColumns = async (res) => {
    await accumulateData(res);
  };

  const onlyAdminsCanSee = (): boolean => {
    let visibilitySetting:boolean = props.visibilityOpts == undefined ? false : props.visibilityOpts ;
    return visibilitySetting;
  };

  const whoCanEdit  = (): boolean => {
    let editSetting: boolean = props.editVisibility == undefined ? false : props.editVisibility ;
    return editSetting;
  };
  
  // Use useEffect to update detailsListData when accumulatedResult changes
  useEffect(() => {
    
    if(!whoCanEdit){
      setWhoCanSee(true);
    }else{
      checkSiteAdmin();
    }

    if(onlyAdminsCanSee()){
      setWhoCanSee(onlyAdminsCanSee());
    }else{
      checkSiteAdmin();
    }
    
    if (accumulatedResult.length > 0 ) {
      setCallLoadColumns(false);
      setDetailsListData(accumulatedResult);
      setIsPanelOpen(false);
    }

    if (!utils.isUndefinedNullOrEmpty(props.query)){
      setCallLoadColumns(true);
      setQueryData(utils.generateOutputArray(props.query));
    }

    // When an admin clicked on the Show Grid button
    if (!utils.isUndefinedNullOrEmpty(detailsListData) && detailsListData.length > 0 ){
      setCallLoadColumns(false);
    }

  }, [accumulatedResult]);

  return (
    <>
      <RowsDataContext.Provider value={{ detailsListData, setDetailsListData }}>
        <div style={{ textAlign: props.buttonAlignment as "center" | "left" | "right" }}>
          {canEdit && 
          <DefaultButton
            style={style}
            className={styles.button}
            iconProps={{ iconName: 'AddToShoppingList', style: { fontSize: `${props.buttonSize}px` } }}
            text={props.buttonTitle}
            onClick={onShowPanel}
          ></DefaultButton>}
        </div>
        {whoCanSee && 
        <ShowPanel webContext={props.webContext} isPanelOpen={isPanelOpen} onPanelDismiss={handlePanelDismiss} onChildDataHandler={handleDataFromLoadColumns} />}
        { callLoadColumns && <DirectGrid webContext={props.webContext} items={queryData} /> }
        { !callLoadColumns && detailsListData.length > 0 && 
        <LoadGrid webContext={props.webContext} columns={columns} items={detailsListData} />}
      </RowsDataContext.Provider>
    </>
  );
};

export default Merge;