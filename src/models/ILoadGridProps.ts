import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ISPColumn } from "./ISPColumn";
export interface ILoadGridProps {
  webContext?: BaseComponentContext;
  columns?: ISPColumn[];
  items?: any;
}


