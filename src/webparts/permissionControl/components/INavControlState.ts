import { INavLinkGroup, INavLink } from "office-ui-fabric-react";

export interface INavControlState {
    SiteName: string;
    results: INavLinkGroup[];
    isloading: boolean;
    showButtonExternalUser: boolean;
}

/*interface results {
    name: string;
    url: string;
    target: string;
    isExpanded: boolean;
    links: results[];
    icon:string
}*/