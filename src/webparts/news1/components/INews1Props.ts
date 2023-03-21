import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface INews1Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
}
