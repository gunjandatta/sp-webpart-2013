import * as React from "react";
import { WebPart } from "gd-sprest-react";
import { Configuration } from "./cfg";
import { DemoWebPart } from "./wp";
import { WebPartCfg } from "./wpCfg";

/**
 * WebPart Demo
 */
export class WebPartDemo {
    // Configuration
    static Configuration = Configuration;

    /**
     * Constructor
     */
    constructor() {
        // Create an instance of the webpart
        new WebPart({
            cfgElementId: "wp-demoCfg",
            displayElement: DemoWebPart,
            editElement: WebPartCfg,
            targetElementId: "wp-demo"
        });
    }
}