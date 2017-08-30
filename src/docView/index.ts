import { WebPart, WebPartSearchCfg } from "gd-sprest-react";
import { Configuration } from "./cfg";
import { DocViewWebPart } from "./wp";
import "./docView.scss";

/**
 * Document View
 */
export class DocView {
    // Configuration
    static Configuration = Configuration;

    /**
     * Constructor
     */
    constructor() {
        // Create an instance of the webpart
        new WebPart({
            cfgElementId: "wp-docViewCfg",
            displayElement: DocViewWebPart,
            editElement: WebPartSearchCfg,
            targetElementId: "wp-docView",
        });
    }
}