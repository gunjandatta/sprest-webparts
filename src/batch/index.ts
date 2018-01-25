import { List } from "gd-sprest";
import { WebParts } from "gd-sprest-react";
import { Configuration } from "./cfg";
import { BatchWebPart } from "./wp";

/**
 * Batch
 */
export class Batch {
    // Configuration
    static Configuration = Configuration;

    /**
     * Constructor
     */
    constructor() {
        // Create an instance of the webpart
        new WebParts.FabricWebPart({
            displayElement: BatchWebPart,
            targetElementId: "wp-batch",
        });
    }
}