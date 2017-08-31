import { WebPart, WebPartSearchCfg } from "gd-sprest-react";
import { Configuration } from "./cfg";
import { ListWebpart } from "./wp";

/**
 * List
 */
export class List {
    // Configuration
    static Configuration = Configuration;

    /**
     * Constructor
     */
    constructor() {
        // Create an instance of the webpart
        new WebPart({
            cfgElementId: "wp-listCfg",
            displayElement: ListWebpart,
            editElement: WebPartSearchCfg,
            targetElementId: "wp-list",
        });
    }
}