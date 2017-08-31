import { DocView } from "./docView";
import { HelloWorld } from "./helloWorld";
import { List } from "./list";
declare var SP;

// Create the Demo WebParts global variable
window["DemoWebParts"] = {
    DocView,
    HelloWorld,
    List
};

// Let SharePoint know this file has loaded
SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("webparts.js")