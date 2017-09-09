import { WebPart, WebPartTabs } from "gd-sprest-react";
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
        new WebPart({
            displayElement: WebPartTabs,
            targetElementId: "wp-tabs",
        });
    }
}