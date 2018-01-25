import { WebParts } from "gd-sprest-react";
import { Configuration } from "./cfg";
import { ListWebpart } from "./wp";
import "./list.scss";

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
        new WebParts.FabricWebPart({
            cfgElementId: "wp-listCfg",
            displayElement: ListWebpart,
            editElement: WebParts.WebPartSearchCfg,
            targetElementId: "wp-list",
        });
    }
}