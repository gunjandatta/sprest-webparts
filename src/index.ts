import { Batch } from "./batch";
import { DocView } from "./docView";
import { Email } from "./email";
import { HelloWorld } from "./helloWorld";
import { List } from "./list";
import { Tabs } from "./tabs";
declare var SP;

// Create the Demo WebParts global variable
window["DemoWebParts"] = {
    Batch,
    DocView,
    Email,
    HelloWorld,
    List,
    Tabs
};

// Let SharePoint know this file has loaded
SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("webparts.js")