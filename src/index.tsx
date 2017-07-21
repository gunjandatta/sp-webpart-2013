import { WebPartDemo } from "./demo";

// Create the global variable
window["Demo"] = {
    // Demo Webpart
    WebPart: WebPartDemo
};

// Let SharePoint know the script has been loaded
window["SP"].SOD.notifyScriptLoadedAndExecuteWaitingJobs("demo.js");