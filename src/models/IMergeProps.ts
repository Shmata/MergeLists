import { BaseComponentContext, IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IMergeProps {
  webContext: BaseComponentContext;
  buttonTitle?: string;
  buttonAlignment?: string ;
  visibilityOption?: string;
  isAdmin?: boolean;
  buttonSize?: number;
  themeVariant?: IReadonlyTheme | undefined;
  visibilityOpts: boolean;
  editVisibility: boolean;
  query:string;
}
