import * as React from 'react';
import Graph from './Graph';
import AppOptions from '../AppOptions';
import { HeroListItem } from './HeroList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { PrimaryButton, ButtonType, MessageBarType} from 'office-ui-fabric-react';
import NotificationMessage, { INotificationParams } from "./NotificationMessage";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import HorizontalSeparator from './HorizontalSeparator';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import Progress from './Progress';
import { Stack, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
//import { access } from 'fs';

//let _confirmationPromptResponse: boolean; // user's reply to 'do you want to continue?'
let _graph: Graph;
let _settings: Office.RoamingSettings;
let _appOptions: AppOptions;
let _notificationMessage: string = "";
let _messageBarTypeWanted: MessageBarType;

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  notificationMessage: string;
  messageBarTypeWanted: MessageBarType;
}

export default class App extends React.Component<AppProps, AppState> {

  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      notificationMessage: "",
      messageBarTypeWanted: MessageBarType.error    
    };
    //this.dialogCallback = this.dialogCallback.bind(this); //
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: 'Ribbon',
          primaryText: 'Achieve more with Office integration'
        },
        {
          icon: 'Unlock',
          primaryText: 'Unlock features and functionality'
        },
        {
          icon: 'Design',
          primaryText: 'Create and visualize like a pro'
        }
      ],
      notificationMessage: _notificationMessage
    });
  }

  click = async () => {

      try {
          let phishingEmailItemId = Office.context.mailbox.item.itemId;
          let phishingEmailSubject = Office.context.mailbox.item.subject;
          _graph = new Graph(Office.context.mailbox);
          await _graph.initializeRESTToken(); // For Outlook REST APIs
          //await _graph.getGraphAccessToken();

          // Test functions
          //_graph.getAttachments(Office.context.mailbox.item.itemId);
          //await _graph.getItemSubject();
          //await _graph.getMIMEMessage(Office.context.mailbox.item.itemId);

          _graph.setEmailTo(_appOptions.notificationEmailAddress);
          _graph.setSubject(_appOptions.defaultNotificationEmailSubject);
          if (_appOptions.notificationForwardAction !== 'forward') {
              _graph.setBody(_appOptions.defaultNotificationEmailBodyForwardAsAttachment);
              await _graph.createItem();
              await _graph.getMIMEMessage(phishingEmailItemId); // this puts it in this._mimecontent
              await _graph.addAttachmentItem(phishingEmailItemId, phishingEmailSubject);
              if (_appOptions.displayEmailBeforeSendingWanted) {
                console.log("Displaying newly created item");
                Office.context.mailbox.displayMessageForm(_graph._createdItemId);
                return;  
              }
              else {
                await _graph.sendItem();
                console.log("Back from SendItem");
              }
          } 
          else {
              _graph.setBody(_appOptions.defaultNotificationEmailBodyForward);
              //await _graph.forwardItem(phishingEmailItemId);
              await _graph.createForwardItem(phishingEmailItemId);
              if (_graph.success) {
                  if (_appOptions.displayEmailBeforeSendingWanted) {
                    console.log("Displaying newly created item");
                    Office.context.mailbox.displayMessageForm(_graph._createdItemId);
                    return;  
                  }
                  else {
                    // Just send it
                    await _graph.updateItem(_graph._createdItemId); // Adds the To: field from above
                    if (_graph.success) {
                      await _graph.sendItem();
                      console.log("Back from SendItem!");
                    }
                  }
              }
          }
          if (_graph.success) {
              console.log("Successfully sent the email!");    
              //
              // Determine if the user wants to delete the original email and if so, delete it since we successfully sent it in.
              //
              if (_appOptions.deleteOriginalEmailWanted === true) {
                  await _graph.deleteItem(phishingEmailItemId);
                  if (!_graph.success) {
                      this._showNotificationMessage('Failed to delete phishing message: ' + _graph.errorMessage, MessageBarType.error);
                      return;
                  }
              }      
              //
              // And finally, let the user know it went ok.
              //
              this._showNotificationMessage('Message successfully sent.', MessageBarType.success);
          }
          else {
              console.log("Failed to send email!");
              this._showNotificationMessage('Failed to send message: ' + _graph.errorMessage, MessageBarType.error);
          }
      }
      catch (e) {
        console.log("Error occurred in App.Click. Reason: ", e);
        this._showNotificationMessage('Error occurred: ' + e, MessageBarType.error);
      }
  };

  /*
  //
  // This is normally used for the confirmation dialog but 
  // we no longer use it.  It's kept here in case a particular
  // customer requires it.
  //
  dialogCallback(myResult:boolean):void {
    console.log("dialogResult: ", myResult);
    _confirmationPromptResponse = myResult;
  }; */

  //
  // This is for all the checkboxes presented.
  //
  _onCheckboxChange = (refName: string, ev: React.FormEvent<HTMLElement>, isChecked: boolean): void => {
    console.log('The option ' + refName + ' has been changed to ' + isChecked, ev.type);    
    if (refName === _appOptions.nameDeleteOriginalEmailWanted) {
      _appOptions.deleteOriginalEmailWanted = isChecked;
    } 
    else if (refName === _appOptions.nameDisplayEmailBeforeSendingWanted) {
      _appOptions.displayEmailBeforeSendingWanted = isChecked;
    } 
    else if (refName === _appOptions.nameSendToMSWanted) {
      _appOptions.sendToMSWanted = isChecked;
    } 
    else if (refName === _appOptions.nameSendToFTCWanted) {
      _appOptions.sendToFTCWanted = isChecked;
    } 
    else if (refName === _appOptions.nameSendToDHSWanted) {
      _appOptions.sendToDHSWanted = isChecked;
    } 
    else {
      console.log("Unknown option: " + refName);
    }    
  }

  //
  // This is for chaning the method by which a forward occurs.
  // We can either forward normally (like DHS or FTC) or forward 
  // as an attachment which some vendors require (like Microsoft).
  //
  _onChangeForwardMethod(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
    console.log("ForwardMethodChanged to " + option, ev);
    _appOptions.notificationForwardAction = option.key;
  }

  //
  // Make sure that there is a valid email address.  
  // For now, this is any string with an @ and at least one dot.
  //
  _getErrorMessageEmailAddress = (value: string): string => {
    let errMsg: string = '';
    if (value.length > 6) {
      if (value.includes('@') && (value.includes('.'))) {
        errMsg = '';
        console.log("Email Address Valid");
        if (_appOptions.notificationEmailAddress !== value) {
          _appOptions.notificationEmailAddress = value;
        }
      }
      else {
        errMsg = 'Invalid email address.';
        console.log("Email Address Invalid");
      }
    }
    return errMsg;
  }

  _onRenderLabelWithLink(props) {
    if (props.name===_appOptions.nameSendToMSWanted) {
      return (
        <span>
            Also send to <Link href="https://docs.microsoft.com/en-us/microsoft-365/security/office-365-security/submit-spam-non-spam-and-phishing-scam-messages-to-microsoft-for-analysis">Microsoft Anti-Phishing</Link>
        </span>
      );
    }
    else if (props.name===_appOptions.nameSendToFTCWanted) {
      return (
        <span>
            Also send to the <Link href="https://www.consumer.ftc.gov/articles/how-recognize-and-avoid-phishing-scams">FTC</Link>
        </span>
      );
    }
    else if (props.name===_appOptions.nameSendToDHSWanted)   {
      return (
        <span>
            Also send to the <Link href="https://www.us-cert.gov/report-phishing">US DHS anti-phishing working group</Link>
        </span>
      );
    }
    else {
      return (
        <span>
          {props.name}
        </span>
      )
    }
  }

  //
  // Force a re-rendering which shows the ErrorMessageBar with the given string.
  //
  _showNotificationMessage(notificationMessage:string, messageBarTypeWanted:MessageBarType) {
    _notificationMessage = notificationMessage;
    _messageBarTypeWanted = messageBarTypeWanted;
    console.log("_showNotificationMessage: " + notificationMessage);
    this.setState({
      listItems: [],
      notificationMessage: _notificationMessage,
      messageBarTypeWanted: _messageBarTypeWanted
    });  
  }

  //
  // This is the callback the ErrorMessageBar uses to get the error message text.
  //
  _getNotificationMessage():INotificationParams {
    try {
      const msg:INotificationParams = {message: _notificationMessage, messageBarTypeWanted: _messageBarTypeWanted}
      //msg.message = _notificationMessage;
      //msg.messageBarTypeWanted = _messageBarTypeWanted;
      _notificationMessage = ""; // Empty the string for the next rendering (when the user closes the ErrorMessageBar)
      console.log("Retrieving notification message: " + msg.message);
      return msg; // {message: _notificationMessage, messageBarTypeWanted: _messageBarTypeWanted};  
    }
    catch (e) {
      console.log("Error occurred in _getNotificationMessage. Reason: ", e);
      return null;
    }
  }

  //
  // This is the callback from the initial dialog asking if the user really wants to send.
  //
  _dialogCallback(continueWanted:boolean) {
    try {
      console.log("ContinueWanted: " + continueWanted);
    }
    catch (e) {
      console.log("Error occurred in _dialogCallback. Reason: ", e)
    }
  }

  render() {
    const {
      title,
      isOfficeInitialized,
    } = this.props;

    if (!isOfficeInitialized) {
      console.log("Office is NOT initialized");
      return (
        <Progress
          title={title}
          logo='assets/logo-filled.png'
          message='Please sideload your addin to see app body.'
        />
      );
    }

    //
    // Load the user's settings.
    // 
    console.log("Office is supposedly initialized");
    _appOptions = new AppOptions;
    _settings = Office.context.roamingSettings; 
    let success = _appOptions.Initialize(_settings);
    if (!success) {      
      console.log("Failed to initialize appOptions"); // use the Notification Error in Commands, MessageBar in TaskPane
    }

    console.log("Rendering main");

    // Confirmation Prompt:
    // import DialogBasicExample from './Prompt'
    // <DialogBasicExample hideDialog={false} isDraggable={true} cb={this._dialogCallback} appOptions={_appOptions}/>        
    // 
    // Header with light background, logo, and welcome message
    // <Header logo='assets/logo-filled.png' title={this.props.title} message='Sperry Software' />
    //
    // HeroList from the sample
    // <HeroList message='Discover what Office Add-ins can do for you today!' items={this.state.listItems}>
    // </HeroList>

    const verticalGapStackTokens: IStackTokens = {
      childrenGap: 10,
      padding: 10
    };

    return (
      <div className='ms-welcome'>      
        <HorizontalSeparator content='Required Options'> </HorizontalSeparator>
        <Stack tokens={verticalGapStackTokens}>
          <TextField 
              className="textFieldEmailAddress"
              label="Submit phishing emails to: " 
              required 
              placeholder="Your IT admin's email address" 
              onGetErrorMessage={this._getErrorMessageEmailAddress}
              validateOnFocusOut={true}
              validateOnLoad={false}
              disabled={true}
              defaultValue={_appOptions.notificationEmailAddress}
              />
          <ChoiceGroup
              className="defaultChoiceGroup"
              defaultSelectedKey={_appOptions.notificationForwardAction}
              options={[
                {
                  key: 'forward',
                  text: 'Forward'
                },
                {
                  key: 'forwardAsAttachment',
                  text: 'Forward As Attachment'
                },
              ]}
              onChange={this._onChangeForwardMethod}
              label="Submission type:"
              disabled={true}
              required={true}
            />
          <HorizontalSeparator content='Other Options'> </HorizontalSeparator>
          <Checkbox label="Delete suspect email after submission" 
                    defaultChecked={_appOptions.deleteOriginalEmailWanted}
                    disabled={true}
                    onChange={this._onCheckboxChange.bind(this, _appOptions.nameDeleteOriginalEmailWanted)} />
          <Checkbox label="Display email before submission" 
                    defaultChecked={_appOptions.displayEmailBeforeSendingWanted}
                    disabled={true}
                    onChange={this._onCheckboxChange.bind(this, _appOptions.nameDisplayEmailBeforeSendingWanted)} />          
        </Stack>
        <HorizontalSeparator content=''> </HorizontalSeparator>
        <NotificationMessage cbGetNotificationMessage={this._getNotificationMessage.bind(this)} />
        <Stack tokens={verticalGapStackTokens}>
          <p className='ms-font-l'>This email will be forwarded to corporate security and should not contain private information.  Are you sure you want to submit the currently selected email?</p>
          <PrimaryButton 
            className='ms-welcome__action' 
            buttonType={ButtonType.hero} 
            iconProps={{ iconName: 'ChevronRight' }} 
            onClick={this.click}>Submit</PrimaryButton>
          <p className='ms-font-s'><a target="_blank" href="http://epsp.pxl.int/corp/Information%20Security/SitePages/Anti-Phishing%20Help.aspx">Access phishing threats and current threats</a></p>  
        </Stack>
      </div>
    );
  }
}

/* 
//
// This was in the original version, right underneath the other checkboxes.
//
          <Checkbox name={_appOptions.nameSendToMSWanted} 
                    defaultChecked={_appOptions.sendToMSWanted}
                    onChange={this._onCheckboxChange.bind(this, _appOptions.nameSendToMSWanted)} 
                    onRenderLabel={this._onRenderLabelWithLink}/>
          <Checkbox name={_appOptions.nameSendToDHSWanted} 
                    defaultChecked={_appOptions.sendToDHSWanted}
                    onChange={this._onCheckboxChange.bind(this, _appOptions.nameSendToDHSWanted)} 
                    onRenderLabel={this._onRenderLabelWithLink}/>
          <Checkbox name={_appOptions.nameSendToFTCWanted} 
                    defaultChecked={_appOptions.sendToFTCWanted}
                    onChange={this._onCheckboxChange.bind(this, _appOptions.nameSendToFTCWanted)} 
                    onRenderLabel={this._onRenderLabelWithLink}/>
*/