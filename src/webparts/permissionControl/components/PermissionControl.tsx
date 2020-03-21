import * as React from 'react';
import styles from './PermissionControl.module.scss';
import { IPermissionControlProps } from './IPermissionControlProps';
import { IPermissionControlState } from './IPermissionControlState';
import {ICustomNavLink, LinkType} from './INavControlProps'
import { escape } from '@microsoft/sp-lodash-subset';
import NavControl from './NavControl';
import GridControl from './GridControl';
import { PnPService } from '../services/PnPService'

export default class PermissionControl extends React.Component<IPermissionControlProps, IPermissionControlState> {
  private _pnpService;
  constructor(props) {
    super(props);
    this._pnpService = new PnPService(this.props.context);
    this.onUrlChange = this.onUrlChange.bind(this);
    this.state = {
      url: this.props.context.pageContext.web.absoluteUrl,
      items:[]
    };
    
  }

public componentDidMount(){
  var custNavLink :ICustomNavLink={
    url: '#',
    customUrl: this.state.url,
    webUrl:this.state.url,
    name:'',
    isUniquePermissions:false,
    linkType:LinkType.web
  }
  this._pnpService.getUsersForWeb(custNavLink).then(grpItems=>{
    if(grpItems){
      this.setState({items:grpItems});
    }
  });
}

onUrlChange = (customNavLink:ICustomNavLink) =>{
  this._pnpService.getUsersForWeb(customNavLink).then(grpItems=>{
    this.setState({
      items:grpItems
    });
  });
  
}

public render(): React.ReactElement<IPermissionControlProps> {
  return (
    <div className={ styles.permissionControl }>
      <div className={styles.rowC}>         
          <NavControl onUrlChange={this.onUrlChange} ctx={this.props.context}></NavControl> 
          <GridControl items={this.state.items} />       
      </div>  
    </div>      
  );
}
}
