import * as React from "react";
import { WebParts } from "gd-sprest-react";
import { Configuration } from "./cfg";

/**
 * Hello World
 */
export class HelloWorld {
    /**
     * Configuration
     */
    static Configuration = Configuration;

    /**
     * Constructor
     */
    constructor() {
        // Create an instance of the webpart
        new WebParts.FabricWebPart({
            cfgElementId: "wp-helloWorldCfg",
            onRenderDisplayElement: this.renderDisplayElement,
            onRenderEditElement: this.renderEditElement,
            targetElementId: "wp-helloWorld"
        });
    }

    /**
     * Methods
     */

    // Method to render the display mode component
    private renderDisplayElement = (targetInfo: WebParts.IFabricWebPartInfo) => {
        // Render the element
        return (
            <div>{"The webpart id is: " + targetInfo.cfg.WebPartId}</div>
        );
    }

    // Method to render the edit mode component
    private renderEditElement = (targetInfo: WebParts.IFabricWebPartInfo) => {
        // Render the element
        return (
            <div>{"The page is in edit mode."}</div>
        );
    }
}