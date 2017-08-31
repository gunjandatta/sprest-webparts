import { DocView } from "./docView";
import { HelloWorld } from "./helloWorld";
declare var SP;

// Create the Demo WebParts global variable
window["DemoWebParts"] = {
    DocView,
    HelloWorld
};

// Let SharePoint know this file has loaded
SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("webparts.js")