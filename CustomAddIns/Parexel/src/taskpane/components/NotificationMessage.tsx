import * as React from 'react';
import { MessageBar, MessageBarType, Link } from 'office-ui-fabric-react';
import AppOptions from '../AppOptions';

export interface INotificationParams {
    message: string;
    messageBarTypeWanted: MessageBarType;
}

interface INotificationMessageProps {    
    cbGetNotificationMessage(): INotificationParams; // callback to get the actual error message
}

interface INotificationMessageState {
    stateNotificationMsg: string; 
    stateMessageBarTypeWanted: MessageBarType
}

export default class NotificationMessage extends React.Component<INotificationMessageProps, INotificationMessageState> {
    public state: INotificationMessageState = {
        stateNotificationMsg: "",
        stateMessageBarTypeWanted: MessageBarType.success
    };

    constructor(props, context) {
        super(props, context);
        this.state = {stateNotificationMsg: "",
                      stateMessageBarTypeWanted: MessageBarType.error};
    }

    private closeWindow = (): void => {
        this.setState({stateNotificationMsg: ""});
    }

    public render() {
        let info = this.props.cbGetNotificationMessage();
        let appOptions = new AppOptions;
        let currentThreatsLink = appOptions.defaultCurrentThreatsLink;
        this.state.stateNotificationMsg = info.message;
        this.state.stateMessageBarTypeWanted = info.messageBarTypeWanted;
        if (info.message !== "") {
            console.log("NotificationMessageBar rendering: ", this.state.stateNotificationMsg);
            let msg = this.state.stateNotificationMsg;
            let theMessageBarType = this.state.stateMessageBarTypeWanted;
            if (theMessageBarType==MessageBarType.success) {
                return <MessageBar 
                            messageBarType={theMessageBarType} 
                            isMultiline={false} 
                            onDismiss={this.closeWindow} 
                            dismissButtonAriaLabel="Close">
                            {msg}
                        </MessageBar>;
            } else {
                return <MessageBar 
                            messageBarType={theMessageBarType} 
                            isMultiline={false} 
                            onDismiss={this.closeWindow} 
                            dismissButtonAriaLabel="Close">
                            {msg}
                            <Link href={currentThreatsLink} target="_blank">
                            See our website for details.
                            </Link>
                        </MessageBar>;
            }
        }
        else {
            console.log("NotificationMessageBar rendering hidden");
            return <div></div>
        }
    }
};

