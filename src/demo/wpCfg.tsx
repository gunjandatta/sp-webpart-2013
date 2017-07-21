import * as React from "react";
import { WebPartConfigurationPanel } from "gd-sprest-react";

/**
 * WebPart Configuration
 */
export class WebPartCfg extends WebPartConfigurationPanel {
    // Method to render the webpart configuration panel
    onRenderContents = (cfg) => {
        return (
            <p>This is where your custom edit interface goes.</p>
        );
    }
}