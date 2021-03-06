import { Helper, SPTypes } from "gd-sprest";

/**
 * Test Configuration
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        /** Test List */
        {
            CustomFields: [
                {
                    name: "TestBoolean",
                    schemaXml: '<Field ID="{E6C387B9-AA16-4115-B57F-601720F9D85B}" Name="TestBoolean" StaticName="TestBoolean" DisplayName="Boolean" Type="Boolean">' +
                    '<Default>0</Default>' +
                    '</Field>'
                },
                {
                    name: "TestChoice",
                    schemaXml: '<Field ID="{8B6EB335-3D5C-42B5-A2DB-601720E8A0BC}" Name="TestChoice" StaticName="TestChoice" DisplayName="Choice" Type="Choice">' +
                    '<Default>Choice 3</Default>' +
                    '<CHOICES>' +
                    '<CHOICE>Choice 1</CHOICE>' +
                    '<CHOICE>Choice 2</CHOICE>' +
                    '<CHOICE>Choice 3</CHOICE>' +
                    '<CHOICE>Choice 4</CHOICE>' +
                    '<CHOICE>Choice 5</CHOICE>' +
                    '</CHOICES>' +
                    '</Field>'
                },
                {
                    name: "TestComments",
                    schemaXml: '<Field ID="{0E11F904-4DA2-48E1-B45B-601720923498}" Name="TestComments" StaticName="TestComments" DisplayName="Comments" Type="Note" AppendOnly="TRUE" />'
                },
                {
                    name: "TestDate",
                    schemaXml: '<Field ID="{5BF47BE2-2697-47C1-B6FE-6017207B221A}" Name="TestDate" StaticName="TestDate" DisplayName="Date Only" Type="DateTime" Format="DateOnly" />'
                },
                {
                    name: "TestDateTime",
                    schemaXml: '<Field ID="{0F804508-A8F4-4DE6-9319-601720CE5294}" Name="TestDateTime" StaticName="TestDateTime" DisplayName="Date/Time" Type="DateTime" />'
                },
                {
                    name: "TestLookup",
                    schemaXml: '<Field ID="{ACF5F7EE-629A-452B-8381-60172088E176}" Name="TestLookup" StaticName="TestLookup" DisplayName="Lookup" Type="Lookup" List="SPReact" ShowField="Title" />'
                },
                {
                    name: "TestMultiChoice",
                    schemaXml: '<Field ID="{22AFA098-4B62-4236-8C01-6017208DAB49}" Name="TestMultiChoice" StaticName="TestMultiChoice" DisplayName="Multi-Choice" Type="MultiChoice">' +
                    '<Default>Choice 3</Default>' +
                    '<CHOICES>' +
                    '<CHOICE>Choice 1</CHOICE>' +
                    '<CHOICE>Choice 2</CHOICE>' +
                    '<CHOICE>Choice 3</CHOICE>' +
                    '<CHOICE>Choice 4</CHOICE>' +
                    '<CHOICE>Choice 5</CHOICE>' +
                    '</CHOICES>' +
                    '</Field>'
                },
                {
                    name: "TestMultiLookup",
                    schemaXml: '<Field ID="{68465DA3-34DD-4FEA-BE7A-60172019C4FA}" Name="TestMultiLookup" StaticName="TestMultiLookup" DisplayName="Multi-Lookup" Type="LookupMulti" List="SPReact" Mult="TRUE" ShowField="Title" />'
                },
                {
                    name: "TestMultiUser",
                    schemaXml: '<Field ID="{35C91E16-6C53-4202-B4AA-60172082983A}" Name="TestMultiUser" StaticName="TestMultiUser" DisplayName="Multi-User" Type="User" Mult="TRUE" UserSelectionMode="0" UserSelectionScope="0" />'
                },
                {
                    name: "TestNote",
                    schemaXml: '<Field ID="{0E11F904-4DA2-48E1-B45B-601720977191}" Name="TestNote" StaticName="TestNote" DisplayName="Note" Type="Note" />'
                },
                {
                    name: "TestNumberDecimal",
                    schemaXml: '<Field ID="{8EABA3DF-D439-4C78-B6E9-601720F7C222}" Name="TestNumberDecimal" StaticName="TestNumberDecimal" DisplayName="Decimal" Type="Number" />'
                },
                {
                    name: "TestNumberInteger",
                    schemaXml: '<Field ID="{02CD9CA9-2E41-42B1-B487-6017208731FD}" Name="TestNumberInteger" StaticName="TestNumberInteger" DisplayName="Integer" Type="Number" />'
                },
                {
                    name: "TestUrl",
                    schemaXml: '<Field ID="{9983709F-C54C-4816-AC2C-601720A0553B}" Name="TestUrl" StaticName="TestUrl" DisplayName="Url" Type="URL" />'
                },
                {
                    name: "TestUser",
                    schemaXml: '<Field ID="{041F5349-6D87-4DF8-8A7A-6017206F6F44}" Name="TestUser" StaticName="TestUser" DisplayName="User" Type="User" UserSelectionMode="0" UserSelectionScope="0" />'
                },
            ],
            ListInformation: {
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                Title: "SPReact"
            },
            ViewInformation: [
                {
                    ViewFields: [
                        "LinkTitle", "TestBoolean", "TestChoice", "TestDate", "TestDateTime",
                        "TestLookup", "TestMultiChoice", "TestMultiLookup", "TestMultiUser",
                        "TestNote", "TestNumberDecimal", "TestNumberInteger", "TestUrl", "TestUser"
                    ],
                    ViewName: "All Items"
                }
            ]
        }
    ],

    WebPartCfg: [
        {
            FileName: "wp_list.webpart",
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
                <property name="Title" type="string">List Item Form</property>
                <property name="Description" type="string">The list webpart.</property>
                <property name="ChromeType" type="chrometype">TitleOnly</property>
                <property name="Content" type="string">
                    &lt;script type="text/javascript" src="/sites/dev/siteassets/webparts.js"&gt;&lt;/script&gt;
                    &lt;div id="wp-list"&gt;&lt;/div&gt;
                    &lt;div id="wp-listCfg" style="display:none"&gt;&lt;/div&gt;
                    &lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { new DemoWebParts.List(); }, 'webparts.js');&lt;/script&gt;
                </property>
            </properties>
        </data>
    </webPart>
</webParts>`
        }
    ]
});