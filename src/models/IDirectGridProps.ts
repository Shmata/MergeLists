import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ISPColumn } from "./ISPColumn";
export interface IDirectGridProps {
  webContext?: BaseComponentContext;
  columns?: ISPColumn[];
  items?: any;
}


