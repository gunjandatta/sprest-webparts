import { DocView } from "./docView";
import { Email } from "./email";
import { HelloWorld } from "./helloWorld";
import { List } from "./list";
declare var SP;

// Create the Demo WebParts global variable
window["DemoWebParts"] = {
    DocView,
    Email,
    HelloWorld,
    List
};

// Let SharePoint know this file has loaded
SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("webparts.js")