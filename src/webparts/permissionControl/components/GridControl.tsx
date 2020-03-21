import * as React from 'react';
import { IGridProps, IGridState,IDetailsListItem } from './IGridControlProps';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { PrimaryButton, Modal, DefaultButton } from 'office-ui-fabric-react';
import styles from './GridControl.module.scss';
import { graph, Group } from '@pnp/graph';
import { sp, Web} from '@pnp/sp';
import { User } from '@microsoft/microsoft-graph-types'
export default class GridControl extends React.Component<IGridProps, IGridState> {
  private _selection: Selection;
  private _allItems: Array<IDetailsListItem>;
  private _columns: IColumn[];
  private _userColumns :IColumn[];

  constructor(props: IGridProps, state: IGridState) {
    super(props);  
    
    // Populate with items for demos.
    this._allItems = new Array<IDetailsListItem>();

    this._columns = [
      { key: 'column1', name: 'User/Group', fieldName: 'name', minWidth: 100,
       maxWidth: 200, isResizable: true, onColumnClick: this._onColumnClick },
      { key: 'column2', name: 'SP Group', fieldName: 'spgroup', minWidth: 100,
       maxWidth: 200, isResizable: true, onColumnClick: this._onColumnClick },
      { key: 'column3', name: 'Permission', fieldName: 'permission', minWidth: 100,
       maxWidth: 200, isResizable: true },
      {
        key: 'column4', name: 'Members', fieldName: 'members', minWidth: 100,
        maxWidth: 200, isResizable: true,
        onRender: (item) => {
          return <PrimaryButton onClick={() => this.onMembersButtonClick(item, event)} >View Members</PrimaryButton>;
        }
      },
      {
        key: 'column4', name: 'EARS', fieldName: 'ears', minWidth: 100, 
        maxWidth: 200, isResizable: true,
        onRender: (item) => {
          if(item.name!=""){
           return <PrimaryButton onClick={() => this.onEARSButtonClick(item, event)} >View in EARS</PrimaryButton>;
          }
        }
      },
    ];

    this._userColumns = [
      {
        key: 'Name',
        name: 'Name',
        fieldName: 'displayName',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      }  ,
      {
        key: 'Email',
        name: 'Email',
        fieldName: 'mail',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      } ,
      {
        key: 'JobTitle',
        name: 'Job Title',
        fieldName: 'jobTitle',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      }  

    ];

    this.state = {
     // url:this.props.url,
      items: this.props.items,
      columns: this._columns ,
      showModal:false,
      grpMembers:[],
      grpName:""
    };
  }

  onMembersButtonClick = (item, event)=>{
    if(item.name==""){
      sp.web.siteGroups.getByName(item.spgroup).users.get().then((users) =>{

        var members = users.map(u=>{
          var member = {
          displayName:u.Title,
          mail:u.Email
          };
          return member;
        });


        this.setState({
          grpMembers:members,
          grpName:item.spgroup,
          showModal: true
        });
      })
    }
    else{
      let gid = item.userandgroup;
      gid = gid.substr(gid.lastIndexOf("|") + 1, gid.length - gid.lastIndexOf("|"))
      graph.groups.getById(gid).members.get().then(members => {    
        this.setState({
          grpMembers:members,
          grpName:item.name,
          showModal: true
        })  
      }); 
    }
    event.preventDefault(); 

  }

  _closeModal =()=>{
    this.setState({showModal:false});
  }

  onEARSButtonClick = (item, event) =>{

  }  

  public componentDidMount(){
    this.setState({
      items:this.props.items
    });
    this._allItems=this.props.items;
  }

  public componentWillReceiveProps(props) {
    const grpItems  = this.props.items;
    if (props.items !== grpItems) {     
      this.setState({items:props.items});
      this._allItems=props.items;
    }
  }
 
  public render(): React.ReactElement<IGridProps> {   // const { items } = this.state;
 
    return (
      <div>
        <div className={styles.searchBox}>
          <TextField  label="Filter by User/Group:" onChange={this._onFilter.bind(this)}/>
        </div>       
        <DetailsList
          items={this.state.items}
          columns={this._columns}
          selectionMode={SelectionMode.none}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
       
        />
        <Modal
          titleAriaId={'Members of Group'}
          subtitleAriaId={''}
          isOpen={this.state.showModal}
          onDismiss={this._closeModal}
          isBlocking={false}
          containerClassName={styles.modalContainer}
         
        >
          <div className={styles.header}>
           Group Name : {this.state.grpName}
          </div>
          <div className={styles.close}>
             <DefaultButton onClick={this._closeModal} text="Close" />
          </div>      
          <DetailsList
            items={this.state.grpMembers}
            columns={this._userColumns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}                
          />           
         
        </Modal>        
      </div>
    );
  }
 


  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  };

  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems
    });
  };
 
}

