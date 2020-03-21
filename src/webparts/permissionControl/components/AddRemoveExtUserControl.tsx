import * as React from 'react';
import { PrimaryButton, DefaultButton, Modal,
   TextField, IDropdownOption, DropdownMenuItemType,Dropdown } from 'office-ui-fabric-react';
import styles from './AddRemoveExtUserControl.module.scss';
import { DisplayMode } from '@microsoft/sp-core-library';
import { AzureService } from '../services/AzureServices';
import { PnPService } from '../services/PnPService';
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';

interface IAddRemoveUserState {
  showModal: boolean;
  addButtonClicked: boolean;
  emails:string;
  userGroup:string;
  emailErrorMessage:string;
  userGroupErrorMessage:string;
  failureMessage:string;
  successMessage:string;
  displaySpinner:boolean;
  disableButton:boolean;
}

interface IAddRemoveUserProps { 
  ctx: BaseWebPartContext;
}

// const dropdownStyles: Partial<IDropdownStyles> = {
//   dropdown: { paddingTop: '50px' }
// };

const options: IDropdownOption[] = [
  { key: 'externalSecurityGroup', text: 'External Security Groups', itemType: DropdownMenuItemType.Header },
  { key: 'contributor', text: 'SCJExternalContributor' },
  { key: 'reader', text: 'SCJExternalReader' }
];


export default class AddRemoveExtUserControl extends React.Component<IAddRemoveUserProps,IAddRemoveUserState> {
  private _azureService;
  private _pnpService;
  constructor(props) {
    super(props);
    this._azureService = new AzureService();
    this._pnpService = new PnPService(this.props.ctx);
    this.state = {
       showModal: false,
       addButtonClicked:false,
       emails:"",
       userGroup:"",
       emailErrorMessage:"",
       userGroupErrorMessage:"",
       failureMessage:"",
       successMessage:"",
       displaySpinner:false,
       disableButton:false

       };
  }

  private _onClick = (event) => {
    this.setState(
      {
        showModal: true
      }
    );
  }
 private  _closeModal = (): void => {  
  this._clearStateonExit();
  this.setState({ showModal: false });

  }

  private _getErrorMessage = (value: string): string => {
    this.setState({emailErrorMessage:""});
   // if(this.state.addButtonClicked){
      this.setState({emails:value.trim()});
      return value.trim.length === 0 ? '' : `you have to provide an email address.`;
   // }    
  };

  private _validateFields() : boolean{
    var isFormValid :boolean = true;
    if(this.state.emails == ""){
      this.setState({emailErrorMessage:"You have to provide atleast one email address."});
      isFormValid=false;
    }
    else{
      var emailArray = this.state.emails.split(',');
      var invalidEmailAddress="";
      emailArray.forEach((email, index) => {
        //validate if a valid email address
        const pattern = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
        const result = pattern.test(email.trim());
        if(result != true){
          invalidEmailAddress += index != length-1 ? `${email} ; ` : email;          
        }
        if(invalidEmailAddress != ""){
          this.setState({
            emailErrorMessage:
            "Following Email address are not valid, please remove them and try again: "+invalidEmailAddress
          });

          isFormValid=false;
        }       
      });

    }
    if(this.state.userGroup == ""){
      this.setState({userGroupErrorMessage:"You have to select one user group."});
      isFormValid=false;
    }

    return isFormValid;

  }
  private _clearStateonExit(){

    this.setState({
      userGroupErrorMessage:"",
      userGroup:"",
      emailErrorMessage:"",
      emails:"",
      failureMessage:"",
      successMessage:"",
      displaySpinner:false,
      disableButton:false
  
    });
  }

  private _clearStateonAddButtonClick(){


    this.setState({
      userGroupErrorMessage:"",     
      emailErrorMessage:"",     
      failureMessage:"",
      successMessage:"" 
    });
  }

  private _removeUserHandler = ():void => {

  }

  private _addButtonClickHandler = ():void =>{
    this._clearStateonAddButtonClick();
    this.setState({
      addButtonClicked:true
    }); 

    if(this._validateFields()){
      if(this.state.emailErrorMessage=="" && this.state.userGroupErrorMessage==""
                     && this.state.emails !="" && this.state.userGroup != ""){

          this.setState({disableButton:true,displaySpinner:true});
          try{
              this._azureService
              .AddGroup1(this.props.ctx.pageContext.web.absoluteUrl,this.state.emails,this.state.userGroup, this.props.ctx)
              .then( response =>{
                  var usersToBeAdded = JSON.parse(response).UsersToBeAdded;
                  var groupId = JSON.parse(response).GroupId;
                  var invalidUsers = JSON.parse(response).InvalidUsers;
                  this._pnpService.addExternalUsers(usersToBeAdded,groupId,this.props.ctx.pageContext.web.absoluteUrl)
                  .then(()=>{
                    console.log(invalidUsers);console.log(usersToBeAdded);
                    this.setState({displaySpinner:false,disableButton:false});
                    this.setState({                 
                      failureMessage:"Following users cannot be added because their domains are not allowed: " + invalidUsers,
                      successMessage:"Following users were addded succesfully:" + usersToBeAdded
                    }); 
                  })
                  .catch(error=>{
                    this.setState({displaySpinner:false,disableButton:false});
                    this.setState({                 
                      failureMessage:"Error occurred, please try again. Error Message: " + error,
                      successMessage:""
                    }); 
                  })
              .catch(e=>{
                this.setState({displaySpinner:false,disableButton:false});
                this.setState({                 
                  failureMessage:"Error occurred, please try again. Error Message: " + e,
                  successMessage:""
                }); 
              })
              });
            }
        catch(exception){
          this.setState({displaySpinner:false,disableButton:false});
                this.setState({                 
                  failureMessage:"Error occurred, please try again. Error Message: " + exception,
                  successMessage:""
                }); 
        }

      }
    }
  }

  private _userGroupSelectHandler = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) =>{

    console.log(item.text);
    this.setState({userGroupErrorMessage:""})
    this.setState({userGroup:item.text});
  }

  public render(): React.ReactElement<IAddRemoveUserProps> {

    return (
      <div>
        <PrimaryButton onClick={this._onClick} text='Add / Remove External Users'></PrimaryButton>
        <Modal
          titleAriaId={""}
          subtitleAriaId={""}
          isOpen={this.state.showModal}
          onDismiss={this._closeModal}
          isBlocking={false}
          containerClassName={styles.ModalContainer}
          
        >
          <div className={styles.headerRow}>
            <div >
              Add/Remove External Users
            </div>
            <div className={styles.close}>
             <DefaultButton onClick={this._closeModal} text="Close" />
            </div>     
          </div>

          <div className={styles.innerContainer}>
            <div>
              <TextField className={styles.item}         
            
               label="Enter External User Email Address. Use comma to seperate the email addresses"
               multiline rows={3} 
               required  
               onGetErrorMessage={this._getErrorMessage}
               errorMessage={this.state.emailErrorMessage} /> 

              <Dropdown className={styles.item} 
               placeholder="Please Select a Security Group" options={options}
               required  errorMessage={this.state.userGroupErrorMessage}
               onChange={this._userGroupSelectHandler} /> 

              <div className={styles.buttonContainer}>               
                  <PrimaryButton className={styles.item} text='Add User' disabled={this.state.disableButton}
                  onClick={this._addButtonClickHandler} ></PrimaryButton> 
                  <PrimaryButton className={styles.item} text='Remove User' disabled={this.state.disableButton}
                  ></PrimaryButton>
              </div>

            </div>  
            {this.state.displaySpinner && 
            <Spinner label="Addin user(s), please wait..." 
            ariaLive="assertive" labelPosition="left" />
             }
            <div className={styles.messageSuccess}>
              {this.state.successMessage}
            </div>
            <div className={styles.messageFailure}>
              {this.state.failureMessage}
            </div>

          </div>
         

        </Modal>
      </div>
    );
  }
}