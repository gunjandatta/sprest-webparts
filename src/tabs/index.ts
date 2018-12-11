import { WebParts } from "gd-sprest-react";
import { Configuration } from "./cfg";

/**
 * Tabs
 */
export class Tabs {
    // Configuration
    static Configuration = Configuration;

    /**
     * Constructor
     */
    constructor() {
        // Create an instance of the webpart
        WebParts.FabricWebPart({
            displayElement: WebParts.WebPartTabs,
            targetElementId: "wp-tabs",
        });
    }
}