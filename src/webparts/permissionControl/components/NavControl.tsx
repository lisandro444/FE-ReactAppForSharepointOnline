import * as React from 'react';
import { INavControlProps, ICustomNavLink } from './INavControlProps';
import { INavControlState } from './INavControlState';
import { Nav, INavLinkGroup, INavLink } from 'office-ui-fabric-react/lib/Nav';
import AddRemoveExtUserControl from './AddRemoveExtUserControl';
import { sp, Web, List, Lists } from '@pnp/sp';
import { PnPService } from '../services/PnPService';
import navStyles from './NavControl.module.scss';
import { AzureService } from '../services/AzureServices';
export default class NavControl extends React.Component<INavControlProps, INavControlState> {

  private _navService;
  private _azureService;

  constructor(props) {
    super(props);
    this._navService = new PnPService(this.props.ctx);
    this._azureService = new AzureService();
    this.state = {
      SiteName: "",
      results: [],
      isloading: false,
      showButtonExternalUser:false
    };

   // this._onLinkClick = this._onLinkClick.bind(this);
  }

  public componentDidMount() {
   
    this._navService.getWeb(this.props.ctx.pageContext.web.absoluteUrl).then(data => {        
        this.setState({ results: data});       
    }); 

    this._azureService.sharingCapabilityAzureFunction(this.props.ctx)
    .then(response =>{
      this.setState({ showButtonExternalUser:JSON.parse(response).DisplayAddRemoveButton });  
     }
    );

  }

  public _onLinkClick(ev: React.MouseEvent<HTMLElement, MouseEvent >, item?: ICustomNavLink) {
    ev.preventDefault();
    this.props.onUrlChange(item);     
  }

  
  public render(): React.ReactElement<INavControlProps> {
    
    return (      
      <div className={navStyles.navControl}>
         <div className={navStyles.navOverride} >           
          {this.state.showButtonExternalUser && <AddRemoveExtUserControl ctx={this.props.ctx}></AddRemoveExtUserControl>}
          {!this.state.showButtonExternalUser && <div className={navStyles.emptyButtonSpace}></div>}
          <Nav
            onLinkClick={this._onLinkClick.bind(this)}
            expandButtonAriaLabel="Expand or collapse"
            styles={{
              root: {
                width: 200,               
                boxSizing: 'border-box',
                //border: '1px solid #eee',
                overflowY: 'auto'
              }}             
            }
            groups={this.state.results}
          />
      </div>
      </div>
    );
  }
}

