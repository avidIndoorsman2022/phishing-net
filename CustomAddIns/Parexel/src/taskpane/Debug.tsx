import { ApplicationInsights, ITelemetryItem, SeverityLevel } from '@microsoft/applicationinsights-web'

class Debug {

    public APP_VER:string = "1.2.7599.61107"; // from the manifest
    private APP_NAME:string = "Phishing Net";
    private APP_NAME_NO_SPACES:string = "PhishingNet";
    private BUILT_FOR_CUSTOMER:string = "Parexel"; // normally blank=retail

    private isInitialized:boolean=false;

    private appInsights = new ApplicationInsights({ config: {
        instrumentationKey: "08b3400d-29a6-47a3-83b4-2b8edf6c16e1", // sperry365\dev-addins
        disableAjaxTracking: true,
        maxAjaxCallsPerView: 20, // reduce default Ajax dependency tracking from 500
        disableFetchTracking: false,
        loggingLevelTelemetry: 0, // disables tracing insights internal logging
        maxBatchInterval: 0, // normally 15000, time between updates to App Insights
        namePrefix: this.APP_NAME_NO_SPACES
      } });

    private Initialize() {
          this.appInsights.loadAppInsights();
          this.appInsights.context.application.ver = this.APP_VER;
          this.appInsights.addTelemetryInitializer((envelope: ITelemetryItem) => {
              if (envelope && envelope.data) {
                  envelope.data['app_name'] = this.APP_NAME;
                  envelope.data['customer'] = this.BUILT_FOR_CUSTOMER; // 'Retail'
              }
          });
          this.appInsights.trackPageView(); // Manually call trackPageView to establish the current user/session/pageview          
          this.isInitialized = true;
          console.log("Debug is initialized")
    };

    //
    // Writes any available information about the environment.
    // It's designed to be called after initializing for the first time.
    // We assume that Office is initialized by this point.
    //
    public DebugInfo() {
        this.Log("DebugInfo", "App: " + this.APP_NAME_NO_SPACES);
        this.Log("DebugInfo", "Version: " + this.APP_VER);
        if (this.BUILT_FOR_CUSTOMER != "") {
            this.Log("DebugInfo", "BuiltFor: " + this.BUILT_FOR_CUSTOMER);
        } else {
            this.Log("DebugInfo", "BuiltFor: Retail");
        }

        try {
            this.Log("DebugInfo", "UserName: " + Office.context.mailbox.userProfile.displayName);
            this.Log("DebugInfo", "EmailAddress: " + Office.context.mailbox.userProfile.emailAddress);
            this.Log("DebugInfo", "TimeZone: " + Office.context.mailbox.userProfile.timeZone);
    
            this.Log("DebugInfo", "Host: " + Office.context.diagnostics.host);
            this.Log("DebugInfo", "Platform: " + Office.context.diagnostics.platform);
            this.Log("DebugInfo", "Version: " + Office.context.diagnostics.version);
    
            this.Log("DebugInfo", "OWAView: " + Office.context.mailbox.diagnostics.OWAView);
            this.Log("DebugInfo", "HostName: " + Office.context.mailbox.diagnostics.hostName);
            this.Log("DebugInfo", "HostVersion: " + Office.context.mailbox.diagnostics.hostVersion);    
        }
        catch (ex) {
            this.Log("DebugInfo", "Error occurred");
            this.appInsights.trackException(
                {
                    exception: ex,                
                    severityLevel: SeverityLevel.Error
                }
            );    
        }
    }

    public Log(title:string, message:string, anyVariable?:any) {
        if (this.isInitialized===false) {
            this.Initialize();
        }
        if (anyVariable==null) {
            console.log(title + ': ' + message);
        } else {
            console.log(title + ': ' + message, anyVariable);
        }
        this.appInsights.setAuthenticatedUserContext(Office.context.mailbox.userProfile.emailAddress,
                                                     Office.context.mailbox.userProfile.displayName);
        this.appInsights.trackTrace(
            {message:title + ': ' + message,                
             severityLevel: SeverityLevel.Information});
    };

    public LogException(title:string, e:any) {
        if (this.isInitialized===false) {
            this.Initialize();
        }
        this.Log(title, "Error occurred");
        this.appInsights.trackException(
            {
                exception: e,                
                severityLevel: SeverityLevel.Error
            }
        );
    }

};
const debug = new Debug();
export default debug;
