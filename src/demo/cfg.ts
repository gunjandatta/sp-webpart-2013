import { ContextInfo, Helper } from "gd-sprest";
declare var SP;

/**
 * Data Source
 */
export const Configuration = new Helper.SPConfig({
    // Custom Action
    CustomActionCfg: {
        Web: [
            {
                Description: "Custom ribbon dropdown for wiki and webpart pages in edit mode.",
                Group: "Demo",
                Location: "CommandUI.Ribbon",
                Name: "Demo_WebRibbon",
                Title: "Demo - Web Ribbon",
                CommandUIExtension:
                `
<CommandUIExtension>
    <CommandUIDefinitions>
        <CommandUIDefinition Location="Ribbon.WebPartPage.Actions.Controls._children">
            <Button
                Id="DemoAddWebPart"
                Command="DemoAddWebPart"
                Image32by32="/_layouts/15/1033/images/formatmap32x32.png?rev=44"
                Image32by32Left="-443"
                Image32by32Top="-375"
                LabelText="Add Demo"
                Description="Add the demo webpart"
                TemplateAlias="o1"
            />
        </CommandUIDefinition>
    </CommandUIDefinitions>
    <CommandUIHandlers>
        <CommandUIHandler
            Command="DemoAddWebPart"
            CommandAction="javascript:Demo.WebPart.Configuration.addDemoWebPart();"
        />
    </CommandUIHandlers>
</CommandUIExtension>
`
            }
        ]
    },

    // WebPart Configuration
    WebPartCfg: [
        {
            FileName: "dev_wpDemo.webpart",
            Group: "Dev",
            XML: `<?xml version="1.0" encoding="utf-8"?>
<webParts>
    <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
            <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
            <importErrorMessage>$Resources:core,ImportantErrorMessage;</importErrorMessage>
        </metaData>
        <data>
            <properties>
                <property name="Title" type="string">Demo Webpart</property>
                <property name="Description" type="string">Demo webpart from a generated webpart file.</property>
                <property name="ChromeType" type="chrometype">None</property>
                <property name="Content" type="string">
                    &lt;div id="wp-demo"&gt;&lt;/div&gt;
                    &lt;div id="wp-demoCfg" style="display: none;"&gt;&lt;/div&gt;
                    &lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { new Demo.WebPart(); }, "demo.js");&lt;/script&gt;
                </property>
            </properties>
        </data>
    </webPart>
</webParts>`
        }
    ]
});

// Method to add a webpart to the current page
Configuration["addDemoWebPart"] = () => {
    // Get the current context
    let context = SP.ClientContext.get_current();

    // Get the webpart from the current page
    let page = context.get_web().getFileByServerRelativeUrl(ContextInfo.serverRequestPath);
    let wpMgr = page.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);

    // Import the webpart
    let wpDef = wpMgr.importWebPart(`<?xml version="1.0" encoding="utf-8"?>
<webParts>
    <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
            <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
            <importErrorMessage>$Resources:core,ImportantErrorMessage;</importErrorMessage>
        </metaData>
        <data>
            <properties>
                <property name="Title" type="string">Demo Webpart</property>
                <property name="Description" type="string">Demo webpart added by a custom ribbon button.</property>
                <property name="ChromeType" type="chrometype">TitleOnly</property>
                <property name="Content" type="string">
                    &lt;div id="wp-demo"&gt;&lt;/div&gt;
                    &lt;div id="wp-demoCfg" style="display: none;"&gt;&lt;/div&gt;
                    &lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { new Demo.WebPart(); }, "demo.js");&lt;/script&gt;
                </property>
            </properties>
        </data>
    </webPart>
</webParts>`);

    // Get the first webpart zone on the page
    let wpZone: any = document.querySelector("#MSOZone");
    wpZone = wpZone ? wpZone.getAttribute("zoneid") : null;
    if (wpZone) {
        // Get the webpart and add it to the page
        var wp = wpDef.get_webPart();
        wpMgr.addWebPart(wp, wpZone, 0);
        context.load(wp);

        // Execute the request
        context.executeQueryAsync(
            // Success
            () => {
                // Disable the edit page warning
                if (SP && SP.Ribbon && SP.Ribbon.PageState && SP.Ribbon.PageState.PageStateHandler) {
                    SP.Ribbon.PageState.PageStateHandler.ignoreNextUnload = true;
                }

                // Refresh the page
                window.location.href = window.location.href;
            },
            // Error
            (...args) => {
                // Log
                console.error("Error adding the webpart.")
                console.error(args[1].get_message());
            }
        );
    } else {
        // Log
        console.error("Unable to detect a webpart zone on the page.");
    }
}