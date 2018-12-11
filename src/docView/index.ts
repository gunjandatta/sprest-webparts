import { WebParts } from "gd-sprest-react";
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
        WebParts.FabricWebPart({
            cfgElementId: "wp-docViewCfg",
            displayElement: DocViewWebPart,
            editElement: WebParts.WebPartSearchCfg,
            targetElementId: "wp-docView",
        });
    }
}