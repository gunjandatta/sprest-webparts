import * as React from "react";
import { WebPart, IWebPartTargetInfo } from "gd-sprest-react";
import { EmailWebPart } from "./wp";
import {Configuration} from "./cfg";

/**
 * Email
 */
export class Email {
    // Configuration
    static Configuration = Configuration;
    
    /**
     * Constructor
     */
    constructor() {
        // Create an instance of the webpart
        new WebPart({
            displayElement: EmailWebPart,
            targetElementId: "wp-email"
        });
    }
}