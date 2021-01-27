import debug from './Debug';

export default class AppOptions {

    //
    // These are the available options.
    //
    private _notificationEmailAddress: string;
    private _notificationForwardAction: string; // either "forward" or "forwardAsAttachment"
    private _displayEmailBeforeSendingWanted: boolean;
    private _deleteOriginalEmailWanted: boolean;
    private _sendToMSWanted: boolean;
    private _sendToDHSWanted: boolean;
    private _sendToFTCWanted: boolean;

    //
    // Provide out-of-the-box defaults for users.
    //
    private _defaultNotificationEmailAddress:string = "";
    private _defaultNotificationForwardAction:string = "forwardAsAttachment";
    private _defaultDisplayEmailBeforeSendingWanted:boolean = false;
    private _defaultDeleteOriginalEmailWanted:boolean = true;
    private _defaultSendToMSWanted:boolean = false;
    private _defaultSendToDHSWanted:boolean = false;
    private _defaultSendToFTCWanted:boolean = false;
    public defaultCurrentThreatsLink:string = "https://us-cert.cisa.gov/ncas/current-activity"

    public readonly defaultNotificationEmailSubject = "Suspected Phishing Email"
    public readonly defaultNotificationEmailBodyForward = "This forwarded email looks suspicious."
    public readonly defaultNotificationEmailBodyForwardAsAttachment = "The attached email looks suspicious."

    //
    // Provide the names for the Settings calls so there are no mistakes.
    //
    public readonly nameNotificationEmailAddress = "notificationEmailAddress";
    public readonly nameNotificationForwardAction = "notificationForwardAction";
    public readonly nameDisplayEmailBeforeSendingWanted = "displayEmailBeforeSendingWanted";
    public readonly nameDeleteOriginalEmailWanted = "deleteOriginalEmailWanted";
    public readonly nameSendToMSWanted = "sendToMSWanted";
    public readonly nameSendToDHSWanted = "sendToDHSWanted";
    public readonly nameSendToFTCWanted = "sendToFTCWanted";

    //
    // For error handling, any functions that fail leave their error message here to be fetched.
    //
    private _isDirty: boolean;
    private _errMessage: string;

    //
    // get/set properties
    //
    get notificationEmailAddress(): string {
        return this._notificationEmailAddress;        
    }

    set notificationEmailAddress(newNotificationEmailAddress: string) {
        if (this._notificationEmailAddress !== newNotificationEmailAddress) {
            this._notificationEmailAddress = newNotificationEmailAddress;
            this._isDirty = true;
        }
    }

    get notificationForwardAction(): string {
        return this._notificationForwardAction;        
    }

    set notificationForwardAction(newNotificationForwardAction: string) {
        if (this._notificationForwardAction !== newNotificationForwardAction) {
            this._notificationForwardAction = newNotificationForwardAction;
            this._isDirty = true;
        }
    }

    get displayEmailBeforeSendingWanted(): boolean {
        return this._displayEmailBeforeSendingWanted;
    }

    set displayEmailBeforeSendingWanted(newDisplayEmailBeforeSendingWanted: boolean) {
        if (this._displayEmailBeforeSendingWanted !== newDisplayEmailBeforeSendingWanted) {
            this._displayEmailBeforeSendingWanted = newDisplayEmailBeforeSendingWanted;
            this._isDirty = true;    
        }
    }
    get deleteOriginalEmailWanted(): boolean {
        return this._deleteOriginalEmailWanted;
    }

    set deleteOriginalEmailWanted(newDeleteOriginalEmailWanted: boolean) {
        if (this._deleteOriginalEmailWanted !== newDeleteOriginalEmailWanted) {
            this._deleteOriginalEmailWanted = newDeleteOriginalEmailWanted;
            this._isDirty = true;    
        }
    }

    get sendToMSWanted(): boolean {
        return this._sendToMSWanted;
    }

    set sendToMSWanted(newSendToMSWanted: boolean) {
        if (this._sendToMSWanted !== newSendToMSWanted) {
            this._sendToMSWanted = newSendToMSWanted;
            this._isDirty = true;    
        }
    }

    get sendToDHSWanted(): boolean {
        return this._sendToDHSWanted;
    }

    set sendToDHSWanted(newSendToDHSWanted: boolean) {
        if (this._sendToDHSWanted !== newSendToDHSWanted) {
            this._sendToDHSWanted = newSendToDHSWanted;
            this._isDirty = true;    
        }
    }

    get sendToFTCWanted(): boolean {
        return this._sendToFTCWanted;
    }

    set sendToFTCWanted(newSendToFTCWanted: boolean) {
        if (this._sendToFTCWanted !== newSendToFTCWanted) {
            this._sendToFTCWanted = newSendToFTCWanted;
            this._isDirty = true;    
        }
    }

    get isDirty(): boolean {
        return this._isDirty;
    }

    get errorMessage(): string {
        return this._errMessage;
    }

    //
    // Load the settings, providing defaults if necessary.
    //
    public Initialize(settings: Office.RoamingSettings): boolean {
        let theSetting: any;
        let success: boolean;

        try {
            this._errMessage = "";

            this._notificationEmailAddress = this._defaultNotificationEmailAddress;
            this._notificationForwardAction = this._defaultNotificationForwardAction;
            this._displayEmailBeforeSendingWanted = this._defaultDisplayEmailBeforeSendingWanted;
            this._deleteOriginalEmailWanted = this._defaultDeleteOriginalEmailWanted;
            this._sendToMSWanted = this._defaultSendToMSWanted;
            this._sendToDHSWanted = this._defaultSendToDHSWanted;
            this._sendToFTCWanted = this._defaultSendToFTCWanted;

            if (settings!=undefined) {

                //
                // Start by getting the setting for the debug log wanted,
                // then we can get all the other settings.
                //
                theSetting = settings.get(debug.nameDebugLogWanted);
                if (theSetting != undefined) {
                    debug.debugLogWanted = theSetting;
                    debug.Log("AppOptions.Initialize", "DebugLogWanted setting: " + debug.debugLogWanted);
                    debug.ResetIsDirty(); // It's not dirty until the user changes it, not us
                }

                debug.Log("AppOptions.Initialize", "Initializing app options");

                theSetting = settings.get(this.nameNotificationEmailAddress);
                if (theSetting != undefined) {
                    this._notificationEmailAddress = theSetting;
                    debug.Log("AppOptions.Initialize", "Fetching notification email address: " + theSetting);
                }
                else {
                    this._notificationEmailAddress = this._defaultNotificationEmailAddress;
                    debug.Log("AppOptions.Initialize", "Fetching default notification email address: " + this._defaultNotificationEmailAddress);
                }

                theSetting = settings.get(this.nameNotificationForwardAction);
                if (theSetting != undefined) {
                    this._notificationForwardAction = theSetting;
                }
                else {
                    this._defaultNotificationForwardAction = this._defaultNotificationForwardAction;
                }

                theSetting = settings.get(this.nameDisplayEmailBeforeSendingWanted);
                if (theSetting != undefined) {
                    this._displayEmailBeforeSendingWanted = theSetting;
                } else {
                    this._displayEmailBeforeSendingWanted = this._defaultDisplayEmailBeforeSendingWanted;
                }

                theSetting = settings.get(this.nameDeleteOriginalEmailWanted);
                if (theSetting != undefined) {
                    this._deleteOriginalEmailWanted = theSetting;
                } else {
                    this._deleteOriginalEmailWanted = this._defaultDeleteOriginalEmailWanted;
                }

                theSetting = settings.get(this.nameSendToMSWanted);
                if (theSetting != undefined) {
                    this._sendToMSWanted = theSetting;
                } else {
                    this._sendToMSWanted = this._defaultSendToMSWanted;
                }

                theSetting = settings.get(this.nameSendToDHSWanted);
                if (theSetting != undefined) {
                    this._sendToDHSWanted = theSetting;
                } else {
                    this._sendToDHSWanted = this._defaultSendToDHSWanted;
                }

                theSetting = settings.get(this.nameSendToFTCWanted);
                if (theSetting != undefined) {
                    this._sendToFTCWanted = theSetting;
                } else {
                    this._sendToFTCWanted = this._defaultSendToFTCWanted;
                }
            }
            
            debug.Log("AppOptions.Initialize", "AppOption " + this.nameNotificationEmailAddress + ": " + this._notificationEmailAddress);
            debug.Log("AppOptions.Initialize", "AppOption " + this.nameNotificationForwardAction + ": " + this._notificationForwardAction);
            debug.Log("AppOptions.Initialize", "AppOption " + this.nameDisplayEmailBeforeSendingWanted + ": " + this._displayEmailBeforeSendingWanted);
            debug.Log("AppOptions.Initialize", "AppOption " + this.nameDeleteOriginalEmailWanted + ": " + this._deleteOriginalEmailWanted);
            debug.Log("AppOptions.Initialize", "AppOption " + this.nameSendToMSWanted + ": " + this._sendToMSWanted);
            debug.Log("AppOptions.Initialize", "AppOption " + this.nameSendToDHSWanted + ": " + this._sendToDHSWanted);
            debug.Log("AppOptions.Initialize", "AppOption " + this.nameSendToFTCWanted + ": " + this._sendToFTCWanted);

            this._isDirty = false;
            
            success = true;
        }
        catch (error) {
            debug.LogException("AppOptions.Initialize", error);
            this._errMessage = error
            success = false;
        }

        return success;        
    }

    //
    // Save the settings, providing defaults if necessary.
    //
    public Save(settings: Office.RoamingSettings, forceIt:boolean=false): boolean {
        let success: boolean = true;

        try {
            debug.Log("AppOptions.Save", "Saving options: " + this._isDirty);
            if (this._isDirty || forceIt) {
                success = false;
                try {
                    settings.set(this.nameNotificationEmailAddress, this._notificationEmailAddress);
                    settings.set(this.nameNotificationForwardAction, this._notificationForwardAction);
                    settings.set(this.nameDisplayEmailBeforeSendingWanted, this._displayEmailBeforeSendingWanted);
                    settings.set(this.nameDeleteOriginalEmailWanted, this._deleteOriginalEmailWanted);
                    settings.set(this.nameSendToMSWanted, this._sendToMSWanted);
                    settings.set(this.nameSendToDHSWanted, this._sendToDHSWanted);
                    settings.set(this.nameSendToFTCWanted, this._sendToFTCWanted);
                    settings.set(debug.nameDebugLogWanted, debug.debugLogWanted);    
                    settings.saveAsync(asyncResult => {
                        if (asyncResult.status===Office.AsyncResultStatus.Failed) {
                            debug.Log("AppOptions.Save", "Error: SaveSettings failed: " + asyncResult.error.message);
                            this._errMessage = asyncResult.error.message;
                            success = false;
                        }
                        else {
                            success = true;
                            this._isDirty = false;
                            debug.ResetIsDirty();
                            debug.Log("AppOptions.Save", "SaveSettings succeeded");
                        }
                    });            
                }
                catch (error) {
                    debug.LogException("AppOptions.Save", error);
                    this._errMessage = error
                    success = false;
                }
            }
        }
        catch (ex) {
            debug.LogException("AppOptions.Save", ex);
            success = false;
        }
        return success;
    }

}

