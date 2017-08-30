import {DocView} from "./docView";
declare var SP;

// Create the Demo WebParts global variable
window["DemoWebParts"] = {
    DocView
};

// Let SharePoint know this file has loaded
SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("webparts.js")