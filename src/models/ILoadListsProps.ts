import { BaseComponentContext } from "@microsoft/sp-component-base";

export interface ILoadListsProps {
  webContext?: BaseComponentContext;
  siteLists?: string[];
  onChildDataHandler?: (data: any) => void; 
}