import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import * as strings from 'MergeWebPartStrings';
import Merge from './components/Merge';
import { IMergeProps } from '../../models/IMergeProps';
import { SharePointServices } from '../../services/SharePointServices';
import { Utilities } from '../../services/Utilities';
import { PermissionKind, hasPermissions } from '../../models/IPermissions';

export interface IMergeWebPartProps {
  buttonTitle: string;
  buttonAlignment: string;
  buttonSize: number;
  visibilityOpts: boolean;
  editVisibility: boolean;
}

export default class MergeWebPart extends BaseClientSideWebPart<IMergeWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  private spservices: SharePointServices;
  private queryString: string;
  private showSettings:  boolean;
  private utils : Utilities;
  constructor(){
    super();
    this.utils = new Utilities();
  }


  public render(): void {
    const element: React.ReactElement<IMergeProps>  = React.createElement(
      Merge,
      {
        webContext : this.context,
        buttonTitle: this.properties.buttonTitle,
        buttonAlignment: this.properties.buttonAlignment as "center" | "left" | "right",
        buttonSize: this.properties.buttonSize,
        themeVariant: this._themeVariant,
        visibilityOpts: this.properties.visibilityOpts,
        editVisibility: this.properties.editVisibility,
        query: this.queryString
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();
    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
    this.spservices = new SharePointServices(this.context);
    this.spservices.checkIfListExistAndCreate().then();
    try{
      let query = await this.spservices.getLastQuery();
      if(!this.utils.isUndefinedNullOrEmpty(query))
        this.queryString = query;
    }catch (err){
      this.queryString = "";
    }
    
    this.checkSiteAdmin();
    return super.onInit();
  }


  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
  }

  
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected checkSiteAdmin = () =>{
    let apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/EffectiveBasePermissions`;
    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
      },
    }).then(response => response.json()).then(data => {
      const hasFullControlPermission = hasPermissions(data.d.EffectiveBasePermissions, PermissionKind.ManageWeb);
      this.showSettings = hasFullControlPermission;
    }).catch(error => {
      console.log(error);
    });

  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let webpartSettingGroups = [];
    if(this.showSettings){
      webpartSettingGroups.push({
            groupName: strings.BasicGroupName,
            isCollapsed: true,
            groupFields: [
              PropertyPaneTextField('buttonTitle', {
                label: strings.DescriptionFieldLabel
              }),
              PropertyPaneChoiceGroup('buttonAlignment', {
                label: strings.ButtonAlignment,
                options: [
                  { key: 'left', text: strings.ButtonAlignmentLeft },
                  { key: 'center', text: strings.ButtonAlignmentCenter },
                  { key: 'right', text: strings.ButtonAlignmentRight }
                ],
              }),
              PropertyPaneSlider('buttonSize', {
                label: strings.ButtonSize,
                min: 10,
                max: 64
              })
            ],
            
          },
          {
            groupName: strings.VisibilityGroupName,
            isCollapsed: true,
            groupFields: [
              PropertyPaneToggle('visibilityOpts', {
                key: 'visibilityToggle',
                checked: false,
                label: strings.WhoCanSee,
                onText: strings.Admins,
                offText: strings.Everyone
              }),
              PropertyPaneToggle('settingOpts', {
                key: 'settingToggle',
                checked: false,
                label: strings.WhoCanEdit,
                onText: strings.Admins,
                offText: strings.Everyone
              })
            ]
          }            
      );
    }else{
      webpartSettingGroups.push({
        groupName: strings.AccessDenied,
            isCollapsed: true,
            groupFields: []
      });
    }
    return {
      pages: [{
        displayGroupsAsAccordion: true,
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: webpartSettingGroups
      }]
    };
  }
  
}
