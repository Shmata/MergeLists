import { BaseComponentContext } from "@microsoft/sp-component-base";

export interface IShowPanelProps {
  webContext?: BaseComponentContext;
  isPanelOpen: boolean;
  onPanelDismiss: (event?: React.MouseEvent<HTMLElement>) => void;
  onChildDataHandler?: (data: any) => void; 
}
