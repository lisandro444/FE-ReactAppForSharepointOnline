import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { User } from '@microsoft/microsoft-graph-types'

export interface IGridProps {
  //url: string;
  items:IDetailsListItem[];
}

export interface IDetailsListItem {
  key: number;
  name: string; 
  permission:string;
  spgroup:string;
  userandgroup:string;
}

export interface IGridState {
 // url:string;
  items:IDetailsListItem[];
  columns: IColumn[];
  showModal:boolean;
  grpMembers:User[];
  grpName:string;
}