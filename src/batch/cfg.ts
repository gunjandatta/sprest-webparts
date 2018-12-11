import { Helper, SPTypes } from "gd-sprest";

/**
 * Tabs Configuration
 */
export const Configuration = Helper.SPConfig({
    WebPartCfg: [
        {
            FileName: "wp_batch.webpart",
            Group: "Dattabase",
            XML:
            `<?xml version="1.0" encoding="utf-8"?>
<webParts>
    <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
            <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
            <importErrorMessage>$Resources:core,ImportantErrorMessage;</importErrorMessage>
        </metaData>
        <data>
            <properties>
                <property name="Title" type="string">Batch Example</property>
                <property name="Description" type="string">A demo webpart using batch requests.</property>
                <property name="ChromeType" type="chrometype">TitleOnly</property>
                <property name="Content" type="string">
                    &lt;script type="text/javascript" src="/sites/dev/siteassets/webparts.js"&gt;&lt;/script&gt;
                    &lt;div id="wp-batch"&gt;&lt;/div&gt;
                    &lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { new DemoWebParts.Batch(); }, 'webparts.js');&lt;/script&gt;
                </property>
            </properties>
        </data>
    </webPart>
</webParts>`
        }
    ]
});