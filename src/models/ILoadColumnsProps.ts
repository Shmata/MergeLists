import { BaseComponentContext } from "@microsoft/sp-component-base";

export interface ILoadColumnsProps{
  webContext?: BaseComponentContext;
  listItems?: string[];
  onChildDataHandler?: (data: any) => void; 
}