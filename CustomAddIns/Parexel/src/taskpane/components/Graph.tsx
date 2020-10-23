import * as $ from 'jquery';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-client';
import debug from '../Debug';

export default class Graph {

    _accessToken: string;
    _retryGetAccessToken = 0;
    _client: MicrosoftGraph.Client;
    _mailbox: Office.Mailbox;    
    _createdItemId: string;
    _createdItem: Object; // was Office.ComposeMessage but Graph uses lowercase while REST uses uppercase
    _mimeContent: string;
    _sendSuccessful: boolean;
    _sendResponse: string;
    _success: boolean;
    _errorMessage: string;
    _emailTo: string;
    _emailSubject: string;
    _emailBody: string;

    /*
    async getGraphAccessToken() {
        this._success = false;
        try {
            let self = this;
            let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true }); 
            let exchangeResponse = await this.getGraphBootstrapToken(bootstrapToken);
            if (exchangeResponse.claims) {
                // Microsoft Graph requires an additional form of authentication. Have the Office host 
                // get a new token using the Claims string, which tells AAD to prompt the user for all 
                // required forms of authentication.
                let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
                exchangeResponse = await this.getGraphBootstrapToken(mfaBootstrapToken);
            }
            
            if (exchangeResponse.error) {
                // AAD errors are returned to the client with HTTP code 200, so they do not trigger
                // the catch block below.
                this.handleAADErrors(exchangeResponse);
            } 
            else 
            {
                // For debugging:
                // showMessage("ACCESS TOKEN: " + JSON.stringify(exchangeResponse.access_token));
    
                // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
                // in the .fail callback of that call, not in the catch block below.
                self._accessToken = exchangeResponse.access_token;
                self._success = true;
            }
        }
        catch(exception) {
            // The only exceptions caught here are exceptions in your code in the try block
            // and errors returned from the call of `getAccessToken` above.
            if (exception.code) { 
                this.handleClientSideErrors(exception);
            }
            else {
                this._errorMessage = "EXCEPTION: " + JSON.stringify(exception);
            } 
        }
    }
    
    private async getGraphBootstrapToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    
    private handleClientSideErrors(error) {
        switch (error.code) {
    
            case 13001:
                // No one is signed into Office. If the add-in cannot be effectively used when no one 
                // is logged into Office, then the first call of getAccessToken should pass the 
                // `allowSignInPrompt: true` option. Since this sample does that, you should not see
                // this error. 
                this._errorMessage = "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.";  
                break;
            case 13002:
                // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
                // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
                this._errorMessage = "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."; 
                break;
            case 13006:
                // Only seen in Office on the Web.
                this._errorMessage = "Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."; 
                break;
            case 13008:
                // Only seen in Office on the Web.
                this._errorMessage = "Office is still working on the last operation. When it completes, try this operation again."; 
                break;
            case 13010:
                // Only seen in Office on the Web.
                this._errorMessage = "Follow the instructions to change your browser's zone configuration.";
                break;
            default:
                // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
                // to non-SSO sign-in.
                //dialogFallback();
                this._errorMessage = "Error " + error.code + ": " + error.message; // is error.message a thing??
                break;
        }
    }
    
    private handleAADErrors(exchangeResponse) {
        // On rare occasions the bootstrap token is unexpired when Office validates it,
        // but expires by the time it is sent to AAD for exchange. AAD will respond
        // with "The provided value for the 'assertion' is not valid. The assertion has expired."
        // Retry the call of getAccessToken (no more than once). This time Office will return a 
        // new unexpired bootstrap token.
        if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
            &&
            (this._retryGetAccessToken <= 0)) 
        {
            this._retryGetAccessToken++;
            this._accessToken = exchangeResponse.access_token;
            this._success = true;
        }
        else 
        {
            // For all other AAD errors, fallback to non-SSO sign-in.
            // For debugging: 
            // showMessage("AAD ERROR: " + JSON.stringify(exchangeResponse));                   
            //dialogFallback();
            this._errorMessage = "AAD Error: "  + JSON.stringify(exchangeResponse);
        }
    }
    */

    // For Outlook REST APIs
    async getGraphToken() {
        let self = this;
        return await new Promise(async resolve => {
                await this._mailbox.getCallbackTokenAsync({isRest: true}, function(result){
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        // Save the access token.
                        try {
                            self._accessToken = result.value;
                            debug.Log("Graph.getGraphToken", "Successfully received token: ", self._accessToken);
                        }
                        catch (e) {
                            debug.LogException("Graph.getGraphToken", e);
                        }
                    } else {
                        // Handle the error.
                        debug.Log("Graph.getGraphToken", "Failed token! ", result);
                    }
                    resolve();
                })
        });
    }

    // For Outlook REST APIs
    private getAuthenticatedClient(accessToken) {
        // Initialize Graph client
        //let myOptions:MicrosoftGraph.Options;
        const client = MicrosoftGraph.Client.init({
          // Use the provided access token to authenticate
          // requests
          baseUrl: this.getClientURL(),
          authProvider: (done) => {
            done(null, accessToken);
          }
        });              
        debug.Log("App.getAuthenticatedClient", "Successfully authenticated client");
        return client;
    }

    // For Outlook REST APIs
    private getClientURL():string {
        //var getMessageUrl = 'https://graph.microsoft.com/v1.0';
        var getMessageUrl = this._mailbox.restUrl + '/v2.0';
        //var getMessageUrl = 'https://outlook.office.com/api/v2.0';
        return getMessageUrl;
    }

    // For Outlook REST APIs
    async initializeRESTToken() {    
        try {
            await this.getGraphToken();
            this._client = this.getAuthenticatedClient(this._accessToken);
            debug.Log("Graph.initializeRESTToken", "Initialize token completed");    
        }    
        catch (e) {
            debug.LogException("Graph.initializeRESTToken: ", e);
        }
    }


    constructor(mailbox:Office.Mailbox) {
        this._mailbox = mailbox;
    }

    //
    // These two functions can be used to determine if any of the async calls failed and why.
    //
    success():boolean {        
        return this._success;
    }

    errorMessage(): string {
        return this._errorMessage;
    }

    //
    // Simple functions for setting various email properties.
    // These need to be called before calling createItem.
    //
    setEmailTo(newEmailTo:string) {
        this._emailTo = newEmailTo;
    }

    setSubject(newSubject:string) {
        this._emailSubject = newSubject; 
    }

    setBody(newBody:string) {
        this._emailBody = newBody; 
    }
    
    // Make sure the ItemId is REST enabled.    
    private getItemRestId(itemId:string) {
        if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
            // itemId is already REST-formatted.
            return itemId;
        } else {
            // Convert to an item ID for API v2.0.
            return Office.context.mailbox.convertToRestId(
            itemId,
            Office.MailboxEnums.RestVersion.v2_0
            );
        }
    }

    // Useful test function for testing tokens
    async getItemSubject() {
        try {
            let self = this;
            let itemId = self.getItemRestId(Office.context.mailbox.item.itemId);
            self._success = false;
            return await new Promise(async resolve => {        
                // Construct the Graph REST URL to the current item.
                //var getMessageUrl = 'https://graph.microsoft.com/v1.0/me/messages/' + itemId;
                let getMessageUrl = self.getClientURL() + '/me/messages/' + itemId;
                debug.Log("Graph.getItemSubject", "Getting from URL: " + getMessageUrl);
                //let message = await self._client.api("https://outlook.office.com/api/v2.0/me/messages/" + itemId)                
                //                 .get();  
                //var subject = message.Subject;
                //console.log("Subject (Graph client): ", subject);
                //console.log("Message (Graph client): ", message);
                await $.ajax({
                    url: getMessageUrl,
                    dataType: 'application/json',
                    //method: 'GET',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken} //,
                    //data: self.getMessageData() 
                }).always(function(response) {
                    debug.Log("Graph.getItemSubject", "Ajax response: " + response.status, response)
                    if (response.status === 200) {
                        let item = JSON.parse(response.responseText);
                        var subject = item.Subject;
                        debug.Log("Graph.getItemSubject", "Subject (Ajax): " + subject);
                        self._success = true;
                    }
                    else {
                        let errorResponse = response.responseText;
                        self._errorMessage = "Error: " + errorResponse.error.message;
                        debug.Log("Graph.getItemSubject", "Error occurred fetching subject: " + self._errorMessage);
                        self._success = false;
                    }
                });
                resolve();
            });    
        }
        catch (e) {
            debug.LogException("Graph.getItemSubject", e);
        }
    }    

    private getMessageData() {
        const message = JSON.stringify({
            Subject:this._emailSubject,
            Importance:"Normal",
            Body:{
                ContentType:"HTML",
                Content:this._emailBody
            },
            ToRecipients:[
                {
                    EmailAddress:{
                        Address:this._emailTo
                    }
                }
            ]
        })
        return message;
    }

    async createItem() {
        try {
            let self = this;
            self._success = false;
            return await new Promise(async resolve => {        
                //let getMessageUrl = self.getClientURL() + '/me/messages/';
                //console.log("Getting from URL: " + getMessageUrl);
                //let url = this.getClientURL() + '/me/messages';
                //let res = await this._client.api(url)
                //    .post(message);
                //console.log("Successfully created item: ", res);    

                let message = self.getMessageData(); 
                await $.ajax({
                    // For Graph:
                    // url: self.getClientURL() + '/me/messages',
                    // dataType: 'json',
                    // method: 'POST',
                    // headers: { 'Authorization': 'Bearer ' + self._accessToken,
                    //            'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false' },
                    // data: message
                    
                    // For Outlook REST APIs
                    url: self.getClientURL() + '/me/messages',
                    dataType: 'json',
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken,
                               'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false' },
                    data: message
                }).done(function(item){
                    // Message is passed in `item`.
                    //var subject = item.Subject;
                    //console.log("Subject (Ajax): ", subject, item);
                    //console.log("ItemId: ", item.Id)
                    //console.log("Item: ", item)
                    self._createdItemId = self.getItemRestId(item.Id);
                    self._createdItem = item; 
                    //console.log("New item: ", self._createdItem);
                    self._success = true;
                }).fail(function(error){
                    // Handle error.
                    // Note that there are other fields returned:
                    // response.error.code
                    // response.error.innerError.requestId
                    // response.error.innerError.date
                    let response = error.responseJSON;
                    debug.Log("Graph.createItem", "Error occurred creating item: " + response.error.message)
                    self._errorMessage = "Error creating item: " + response.error.message;
                    self._success = false;
                });
                resolve();
            });    
        }
        catch (e) {
            debug.LogException("Graph.createItem", e);            
            this._errorMessage = "Error occurred creating item. Reason: " + e;
            this._success = false;
        }
    }

    
    //
    // Gets the MIME text representing an attachment.
    // 
    async getMIMEMessage(rawItemId:string) {
        try {
            let self = this;
            let itemId = self.getItemRestId(rawItemId);
            self._success = false;
            return await new Promise(async resolve => {        
                // Construct the Graph REST URL to the current item.
                //var getMessageUrl = 'https://graph.microsoft.com/v1.0/me/messages/' + itemId;
                let getMessageUrl = self.getClientURL() + '/me/messages/' + itemId+ '/$value';
                debug.Log("Graph.getMIMEMessage", "Getting from URL: " + getMessageUrl);
                //let message = await self._client.api("https://outlook.office.com/api/v2.0/me/messages/" + itemId)                
                //                 .get();  
                //var subject = message.Subject;
                //console.log("Subject (Graph client): ", subject);
                //console.log("Message (Graph client): ", message);
                await $.ajax({
                    url: getMessageUrl,
                    //dataType: 'application/json',
                    //method: 'GET',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken} //,
                    //data: self.getMessageData() 
                }).done(function(response) {
                    self._mimeContent = response;
                    self._success = true;
                    debug.Log("Graph.getMIMEMessage", "Successfully retrieved MIME content: " + self._mimeContent.length);
                }).fail(function(errorResponse) {
                    let errorMessage = errorResponse.responseJSON;
                    self._errorMessage = "Error: " + errorMessage.error.message;
                    debug.Log("Graph.getMIMEMessage", "Error occurred fetching MIME content: " + errorMessage.error.message, errorResponse);
                    self._success = false;
                });
                resolve();
            });    
        }
        catch (e) {
            debug.Log("Graph.getMIMEMessage", e);
        }
    }

    async getAttachments(itemId) {
        try {
            let self = this;
            self._success = false;
            return await new Promise(async resolve => {        
                //let url = this.getClientURL() + '/me/messages';
                //let res = await this._client.api(url)
                //    .post(message);
                //console.log("Successfully created item: ", res);    

                await $.ajax({
                    url: self.getClientURL() + '/me/messages/' + self.getItemRestId(itemId) + '/attachments',
                    dataType: 'json',
                    method: 'GET',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken}
                }).done(function(listOfAttachments){
                    //var subject = item.Subject;
                    //console.log("Subject (Ajax): ", subject, item);
                    debug.Log("Graph.getAttachments", "List of Attachments: success", listOfAttachments)
                    self._success = true;
                }).fail(function(error){
                    // Handle error.
                    // Note that there are other fields returned:
                    // response.error.code
                    // response.error.innerError.requestId
                    // response.error.innerError.date
                    let response = error.responseJSON;
                    debug.Log("Graph.getAttachments", "Error occurred creating item: " + response.error.message, error)
                    self._errorMessage = "Error retrieving list of attachments: " + response.error.message;
                    self._success = false;
                });
                resolve();
            });    
        }
        catch (e) {
            debug.LogException("Graph.getAttachments", e);            
            this._errorMessage = "Error occurred retrieving list of attachments. Reason: " + e;
            this._success = false;
        }
    }

    async cbAddItemAttachment(result) {
        debug.Log("Graph.cbAddItemAttachment", "Added Item Attachment", result);
    }

    // async jsAddAttachmentItem(itemIdToAttach:string, name:string) {
    //     try {
    //         let self = this;
    //         self._success = false;
    //         return await new Promise(async resolve => {  
    //             let safeName = encodeURI(name);
    //             let contextOptions = { 'asyncContext': { success: self._success, errorMessage: self._errorMessage } }; // The values in asyncContext can be accessed in the callback.
    //             await self._createdItem.addItemAttachmentAsync(itemIdToAttach, 
    //                                                      safeName, 
    //                                                      contextOptions, 
    //                                                      self.cbAddItemAttachment);
    //             resolve();
    //         });    
    //     }
    //     catch (e) {
    //         console.log("Error occurred attaching item.  Reason: ", e);            
    //         this._errorMessage = "Error occurred attaching item. Reason: " + e;
    //         this._success = false;
    //     }
    // }

    //
    // Gets the MIME content of the item to attach then attaches it as a FILE.
    //
    async addAttachmentItem(itemIdToAttach:string, name:string) {
        let self = this;
        try {            
            self._success = false;
            return await new Promise(async resolve => {        
                try {
                //let url = this.getClientURL() + '/me/messages';
                //let res = await this._client.api(url)
                //    .post(message);
                //console.log("Successfully created item: ", res);    
                let safeName = name;                
                debug.Log("Graph.addAttachmentItem", "Attaching: " + name, itemIdToAttach);                
                const attachmentData = JSON.stringify({
                    "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                    //ContentType: null,
                    //"Id": itemIdToAttach, // needs quotes?
                    //"IsInline": "false",
                    //LastModifiedDateTime: modifiedDate,
                    "Name": safeName + ".eml",
                    "ContentBytes": btoa(unescape(encodeURIComponent(self._mimeContent)))
                    });        
                   //let attachmentData = self.getItemAttachmentData(itemIdToAttach, safeName);
                await $.ajax({
                    url: self.getClientURL() + '/me/messages/' + self._createdItemId + '/attachments',
                    dataType: 'json',
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken,
                               'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false' },
                    data: attachmentData
                }).done(function(attachmentResponse){
                    //var subject = item.Subject;
                    //console.log("Subject (Ajax): ", subject, item);
                    debug.Log("Successfully attached", attachmentResponse);
                    self._success = true;
                }).fail(function(error){
                    // Handle error.
                    // Note that there are other fields returned:
                    // response.error.code
                    // response.error.innerError.requestId
                    // response.error.innerError.date
                    let response = error.responseJSON;
                    debug.Log("Graph.addAttachmentItem", "Error occurred attaching item: " + response.error.message)
                    self._errorMessage = "Error attaching attachment: " + response.error.message;
                    self._success = false;
                });
                resolve();
            }
            catch (e) {
                debug.LogException("Graph.addAttachmentItem", e);            
            }
            });    
        }
        catch (e) {
            debug.LogException("Graph.addAttachmentItem", e);            
            this._errorMessage = "Error occurred adding attachments. Reason: " + e;
            this._success = false;
        }
    }

    async forwardItem(itemId:string) {
        try {
            debug.Log("Graph.forwardItem", "Preparing to forward item " + itemId);
            let self = this;
            self._success = false;
            return await new Promise(async resolve => {        
                //let url = this.getClientURL() + '/me/messages/{' + this._createdItemId + '}/forward';
                //let res = await this._client.api(url)
                //    .post(null);
                //console.log("Successfully forwarded item: ", res);    
                const message = JSON.stringify({
                    Comment:this._emailBody,
                    ToRecipients:[
                        {
                            EmailAddress:{
                                Address:this._emailTo
                            }
                        }
                    ]
                });
                await $.ajax({
                    url: self.getClientURL() + '/me/messages/' + self.getItemRestId(itemId) + '/forward',
                    //dataType: 'json',
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken,
                                'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false' },
                    data: message
                }).done(function(response) {
                    self._success = true;
                    debug.Log("Graph.forwardItem", "Successfully sent forwarded email", response);
                }).fail(function(error){
                    // Handle error.
                    // Note that there are other fields returned:
                    // response.error.code
                    // response.error.innerError.requestId
                    // response.error.innerError.date
                    let response = error.responseJSON;
                    debug.Log("Graph.forwardItem", "Error occurred sending forwarded item: "+ response.error.message, error);
                    self._errorMessage = "Error creating item: " + response.error.message;
                    self._success = false;
                });    
                resolve();
            })
        }
        catch (e) {
            debug.LogException("Graph.forwardItem", e);            
            this._errorMessage = "Error occurred sending forward item. Reason: " + e;
            this._success = false;
        }
    }

    async createForwardItem(itemId:string) {
        try {
            debug.Log("Graph.createForwardItem", "Preparing to create forward item " + itemId);
            let self = this;
            self._success = false;
            return await new Promise(async resolve => {        
                //let url = this.getClientURL() + '/me/messages/{' + this._createdItemId + '}/forward';
                //let res = await this._client.api(url)
                //    .post(null);
                //console.log("Successfully forwarded item: ", res);    
                await $.ajax({
                    url: self.getClientURL() + '/me/messages/' + self.getItemRestId(itemId) + '/createforward',
                    dataType: 'json',
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken,
                                'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false' },
                }).done(function(response) {
                    self._success = true;
                    debug.Log("Graph.createForwardItem", "Successfully created forwarded email");
                    self._createdItemId = response.Id;
                    self._createdItem = response; 
                }).fail(function(error){
                    // Handle error.
                    // Note that there are other fields returned:
                    // response.error.code
                    // response.error.innerError.requestId
                    // response.error.innerError.date
                    let response = error.responseJSON;
                    debug.Log("Graph.createForwardItem", "Error occurred creating forward item", error);
                    self._errorMessage = "Error creating forward item: " + response.error.message;
                    self._success = false;
                });    
                resolve();
            })
        }
        catch (e) {
            debug.LogException("Graph.createForwardItem", e);            
            this._errorMessage = "Error occurred sending forward item. Reason: " + e;
            this._success = false;
        }
    }

    async updateItem(itemId:string) {
        try {
            debug.Log("Graph.updateItem", "Preparing to update item " + itemId);
            let self = this;
            self._success = false;
            return await new Promise(async resolve => {        
                //let url = this.getClientURL() + '/me/messages/{' + this._createdItemId + '}/forward';
                //let res = await this._client.api(url)
                //    .post(null);
                //console.log("Successfully forwarded item: ", res);    
                const message = JSON.stringify({
                    ToRecipients:[
                        {
                            EmailAddress:{
                                Address:this._emailTo
                            }
                        }
                    ]
                });
                await $.ajax({
                    url: self.getClientURL() + '/me/messages/' + self.getItemRestId(itemId),
                    dataType: 'json',
                    method: 'PATCH',
                        headers: { 'Authorization': 'Bearer ' + self._accessToken,
                                'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false'},
                    data: message
                }).done(function(response) {
                    self._success = true;
                    debug.Log("Graph.updateItem", "Successfully updated forwarded email", response);
                }).fail(function(error){
                    // Handle error.
                    // Note that there are other fields returned:
                    // response.error.code
                    // response.error.innerError.requestId
                    // response.error.innerError.date
                    let response = error.responseJSON;
                    debug.Log("Graph.updateItem", "Error occurred updating item: " + response.error.message, error)
                    self._errorMessage = "Error updating item: " + response.error.message;
                    self._success = false;
                });    
                resolve();
            })
        }
        catch (e) {
            debug.LogException("Graph.updateItem", e);            
            this._errorMessage = "Error occurred updating forward item. Reason: " + e;
            this._success = false;
        }
    }

    async sendItem() {
        try {
            console.log("Preparing to send item");
            let self = this;
            self._success = false;
            return await new Promise(async resolve => {        
                //let url = this.getClientURL() + '/me/messages/{' + this._createdItemId + '}/send';
                //let res = await this._client.api(url)
                //    .post(null);
                //console.log("Successfully sent item: ", res);    

                await $.ajax({
                    url: self.getClientURL() + '/me/messages/' + self._createdItemId + '/send',
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken},
                }).done(function(response) {
                        self._success = true;
                        console.log("Successfully sent email" + response.status);
                }).fail(function(error){
                        let errorResponse = error.responseText;
                        self._errorMessage = "Error sending email: " + errorResponse;
                        console.log("Error occurred sending email: ", self._errorMessage);
                        self._success = false;
                });
                resolve();
            });    
        }
        catch (e) {
            debug.Log("Error occurred sending item.  Reason: ", e);            
            this._errorMessage = "Error occurred sending item. Reason: " + e;
            this._success = false;
        }
      }

      //
      // Delete is implemented as a move to Deleted Items.
      //
      async deleteItem(itemId:string) {
        try {
            debug.Log("Graph.deleteItem", "Preparing to delete item " + itemId);
            let self = this;
            self._success = false;
            return await new Promise(async resolve => {        
                //let url = this.getClientURL() + '/me/messages/{' + this._createdItemId + '}/forward';
                //let res = await this._client.api(url)
                //    .post(null);
                //console.log("Successfully forwarded item: ", res);    
                const message = JSON.stringify({
                    DestinationId:"DeletedItems"
                });
                await $.ajax({
                    url: self.getClientURL() + '/me/messages/' + self.getItemRestId(itemId) + '/move',
                    dataType: 'json',
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + self._accessToken,
                               'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false'},
                    data: message
                }).done(function(response) {
                    self._success = true;
                    debug.Log("Grapg.deleteItem", "Successfully deleted suspicious email", response);
                }).fail(function(error){
                    // Handle error.
                    // Note that there are other fields returned:
                    // response.error.code
                    // response.error.innerError.requestId
                    // response.error.innerError.date
                    let response = error.responseJSON;
                    debug.Log("Graph.deleteItem", "Error occurred deleting item: ", error)
                    self._errorMessage = "Error deleting item: " + response.error.message;
                    self._success = false;
                });    
                resolve();
            })
        }
        catch (e) {
            debug.Log("Error occurred deleting suspicious item.  Reason: ", e);            
            this._errorMessage = "Error occurred deleting suspicious item. Reason: " + e;
            this._success = false;
        }
      }
}
