import { Demo } from "./demo";

// Create the global variable for the solution
window["Solution"] = {
    // Demo Webpart
    Demo
};

// Let SharePoint know the script has been loaded
window["SP"].SOD.notifyScriptLoadedAndExecuteWaitingJobs("demo.js");