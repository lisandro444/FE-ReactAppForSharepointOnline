import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { INavLink } from "office-ui-fabric-react";

export interface INavControlProps {
  onUrlChange: (customNavLink:ICustomNavLink) => void;
  ctx: BaseWebPartContext;
}

export interface ICustomNavLink extends INavLink {  
  customUrl:string;
  webUrl:string;
  isUniquePermissions:boolean;
  linkType:LinkType
}

export enum LinkType{
  web=0,
  list=1,
  folder=2
}

